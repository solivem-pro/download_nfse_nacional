from __future__ import annotations

import logging
import os
import shutil
import time
import zipfile
from pathlib import Path
from typing import Any, Callable
from unicodedata import normalize

from config.cadastro_db import list_companies
from config.company_workspace import company_control_file, company_period_dir, ensure_company_workspace, ensure_period_workspace
from config.config import DIRETORIOS, ROOT_DIR, Config
from downloader.competencia import NFSeDownloaderCompetencia
from downloader.emissao import NFSeDownloaderEmissao

logger = logging.getLogger(__name__)

EventCallback = Callable[[dict[str, Any]], None]


def prepare_download_companies(selected_codes: list[int | str]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    requested = [str(code) for code in selected_codes]
    companies_by_code = {str(company["cod"]): company for company in list_companies()}

    valid: list[dict[str, Any]] = []
    skipped_expired: list[dict[str, Any]] = []
    missing: list[str] = []

    for code in requested:
        company = companies_by_code.get(code)
        if not company:
            missing.append(code)
            continue

        prepared = {"cod": code, "nome": company["empresa"], "cadastro": company}
        if _is_certificate_expired(company):
            skipped_expired.append(prepared)
        else:
            valid.append(prepared)

    return valid, skipped_expired, missing


def export_company_archives(selected_codes: list[int | str], destination_dir: str | Path) -> dict[str, Any]:
    destination = Path(destination_dir)
    destination.mkdir(parents=True, exist_ok=True)

    config = Config.load(DIRETORIOS["config_json"])
    save_mode_slug = _slug_text(getattr(config, "save_mode", "C\u00f3digo"))
    companies_by_code = {str(company["cod"]): company for company in list_companies()}

    exported: list[dict[str, str]] = []
    missing_archives: list[dict[str, str]] = []
    errors: list[dict[str, str]] = []

    for code in [str(code) for code in selected_codes]:
        company = companies_by_code.get(code)
        if not company:
            missing_archives.append({"cod": code, "empresa": "Empresa nao encontrada"})
            continue

        source = Path(DIRETORIOS["notas"]) / f"{code}.zip"
        if not source.exists():
            missing_archives.append({"cod": code, "empresa": company["empresa"]})
            continue

        try:
            target_name = _export_name_for_company(company, save_mode_slug)
            target = destination / target_name
            if target.resolve(strict=False) == source.resolve(strict=False):
                exported.append({"cod": code, "empresa": company["empresa"], "arquivo": target_name})
                continue
            shutil.copy2(source, target)
            exported.append({"cod": code, "empresa": company["empresa"], "arquivo": target_name})
        except Exception as exc:
            logger.error("Erro ao exportar pacote da empresa %s: %s", code, exc)
            errors.append({"cod": code, "empresa": company["empresa"], "erro": str(exc)})

    return {
        "save_mode": getattr(config, "save_mode", "C\u00f3digo"),
        "exported": exported,
        "missing": missing_archives,
        "errors": errors,
    }


def run_download_batch(
    companies: list[dict[str, Any]],
    year: str,
    month: str,
    on_event: EventCallback | None = None,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    total_companies = len(companies)
    global_documents = 0

    for index, company in enumerate(companies, start=1):
        _emit(
            on_event,
            kind="company_start",
            index=index,
            total=total_companies,
            code=company["cod"],
            name=company["nome"],
            nsu=0,
            documents=global_documents,
            message=f"Iniciando processamento de [{company['cod']}] {company['nome']}",
        )

        result = _download_company(
            company,
            year,
            month,
            company_index=index,
            company_total=total_companies,
            documents_total=global_documents,
            on_event=on_event,
        )
        global_documents += int(result.get("documentos", 0))

        warning = ""
        if result.get("erros", 0) == 0:
            try:
                processed = _process_after_download(company["cod"], company["nome"], year, month)
                if not processed:
                    warning = " Pos-processamento finalizado com aviso."
            except Exception as exc:
                logger.warning("Erro no pos-processamento da empresa %s: %s", company["cod"], exc)
                warning = " Pos-processamento finalizado com aviso."

        if warning:
            result["mensagem"] = f"{result.get('mensagem', '').strip()}{warning}".strip()

        results.append(result)
        _emit(
            on_event,
            kind="company_complete",
            index=index,
            total=total_companies,
            code=company["cod"],
            name=company["nome"],
            nsu=result.get("nsu_final", 0),
            documents=global_documents,
            result=result,
        )

    return results


def _download_company(
    company: dict[str, Any],
    year: str,
    month: str,
    company_index: int,
    company_total: int,
    documents_total: int,
    on_event: EventCallback | None = None,
) -> dict[str, Any]:
    code = company["cod"]
    name = company["nome"]
    cadastro = company["cadastro"]

    try:
        config = Config.load(DIRETORIOS["config_json"])
        consult_mode = getattr(config, "consult_mode", "Emissao")
        cert_rel_path = str(cadastro.get("cert_path", "") or "").strip()
        cert_pass = str(cadastro.get("cert_pass", "") or "")
        cnpj = str(cadastro.get("cnpj", "") or "")

        if _is_certificate_expired(cadastro):
            return _result(code, name, 0, 1, f"Certificado vencido para {name}.")

        if not cert_rel_path:
            return _result(code, name, 0, 1, f"Certificado nao cadastrado para {name}.")
        if not cert_pass:
            return _result(code, name, 0, 1, f"Senha do certificado nao informada para {name}.")

        cert_path = Path(cert_rel_path)
        cert_path = cert_path if cert_path.is_absolute() else Path(ROOT_DIR) / cert_path
        if not cert_path.exists():
            return _result(code, name, 0, 1, f"Arquivo de certificado nao encontrado: {cert_rel_path}")

        workspace_ok, workspace_msg = ensure_company_workspace(code, cnpj, name)
        if not workspace_ok:
            return _result(code, name, 0, 1, workspace_msg)

        period_ok, period_msg = ensure_period_workspace(code, cnpj, name, year, month)
        if not period_ok:
            return _result(code, name, 0, 1, period_msg)

        period_dir = Path(period_msg)
        control_file = company_control_file(code, name)

        company_config = Config.load(DIRETORIOS["config_json"])
        company_config.cert_path = str(cert_path)
        company_config.cert_pass = cert_pass
        company_config.cnpj = cnpj
        company_config.output_dir = str(period_dir)

        downloader = _build_downloader(company_config, consult_mode)
        nsu_current = 0
        docs_total = documents_total

        def write_progress(message: str, log: bool = True) -> None:
            nonlocal nsu_current, docs_total

            if log:
                logger.info("[%s] %s", code, message)
                _emit(on_event, kind="log", code=code, name=name, message=f"[{code}] {message}")

            extracted_nsu = _extract_nsu_from_message(message)
            if extracted_nsu is not None:
                nsu_current = extracted_nsu

            if "XML baixado" in message:
                docs_total += 1
                _emit(
                    on_event,
                    kind="document",
                    code=code,
                    name=name,
                    index=company_index,
                    total=company_total,
                    nsu=nsu_current,
                    documents=docs_total,
                    message=message,
                )
            else:
                _emit(
                    on_event,
                    kind="progress",
                    code=code,
                    name=name,
                    index=company_index,
                    total=company_total,
                    nsu=nsu_current,
                    documents=docs_total,
                    message=message,
                )

        mode_slug = _slug_text(consult_mode)
        if "emiss" in mode_slug:
            downloaded = downloader.run_emissao(
                ano=year,
                mes=month,
                nsu_competencia_file=str(control_file),
                write=write_progress,
            )
        else:
            downloaded = downloader.run_competencia(
                ano_compet=year,
                mes_compet=month,
                nsu_competencia_file=str(control_file),
                write=write_progress,
            )

        pdf_stats = getattr(downloader, "pdf_stats_last_run", None) or {}
        event_stats = getattr(downloader, "event_stats_last_run", None) or {}

        pdf_message = ""
        if getattr(company_config, "download_pdf", False):
            pdf_message = f" | PDFs: {int(pdf_stats.get('sucessos', 0))} ok, {int(pdf_stats.get('falhas', 0))} falhas"

        event_message = (
            f" | Eventos: {int(event_stats.get('salvos', 0))} salvos, "
            f"{int(event_stats.get('sem_evento', 0))} sem evento, "
            f"{int(event_stats.get('falhas', 0))} falhas, "
            f"{int(event_stats.get('cancelamentos', 0))} cancelamentos"
        )

        return {
            "cod": code,
            "empresa": name,
            "documentos": int(downloaded),
            "erros": 0,
            "mensagem": f"Sucesso: {int(downloaded)} documento(s) baixado(s){pdf_message}{event_message}",
            "nsu_final": nsu_current,
        }

    except Exception as exc:
        logger.exception("Erro ao baixar empresa %s", code)
        return _result(code, name, 0, 1, f"Erro durante download de {name}: {exc}")


def _build_downloader(config: Config, consult_mode: str) -> NFSeDownloaderEmissao | NFSeDownloaderCompetencia:
    if "emiss" in _slug_text(consult_mode):
        return NFSeDownloaderEmissao(config)
    return NFSeDownloaderCompetencia(config)


def _process_after_download(cod_company: int | str, company_name: str, year: str, month: str) -> bool:
    period_dir = company_period_dir(cod_company, company_name, year, month)
    xlsm_files = sorted(period_dir.glob("*.xlsm"))
    if not xlsm_files:
        logger.warning("Nenhum arquivo .xlsm encontrado para a empresa %s.", cod_company)
        return False

    report_path = xlsm_files[0]
    macro_ok = _run_excel_macro(report_path, "ImportarTodosXMLs")
    zip_ok = _zip_company_folder(period_dir, Path(DIRETORIOS["notas"]) / f"{cod_company}.zip")
    return macro_ok and zip_ok


def _run_excel_macro(report_path: Path, macro_name: str) -> bool:
    excel = None
    workbook = None

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        logger.warning("pywin32 nao esta disponivel para executar a macro %s.", macro_name)
        return False

    try:
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = 0
        excel.AskToUpdateLinks = 0
        excel.ScreenUpdating = 0
        excel.EnableEvents = 0
        excel.Interactive = 0

        time.sleep(0.5)
        workbook = excel.Workbooks.Open(str(report_path), ReadOnly=False)
        excel.Run(macro_name)
        workbook.Save()
        workbook.Close()
        excel.Quit()
        logger.info("Macro %s executada com sucesso em %s.", macro_name, report_path)
        return True

    except Exception as exc:
        logger.error("Erro ao executar macro %s em %s: %s", macro_name, report_path, exc)
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception:
            pass
        return False

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _zip_company_folder(source_dir: Path, zip_path: Path) -> bool:
    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for root, dirs, files in os.walk(source_dir):
                for dir_name in dirs:
                    dir_path = Path(root) / dir_name
                    relative = dir_path.relative_to(source_dir)
                    archive.writestr(zipfile.ZipInfo(str(relative).replace("\\", "/") + "/"), "")

                for file_name in files:
                    file_path = Path(root) / file_name
                    relative = file_path.relative_to(source_dir)
                    archive.write(file_path, relative)

        logger.info("Pasta da empresa compactada em %s.", zip_path)
        return True
    except Exception as exc:
        logger.error("Erro ao compactar a pasta %s: %s", source_dir, exc)
        return False


def _result(code: str, name: str, documents: int, errors: int, message: str) -> dict[str, Any]:
    return {
        "cod": code,
        "empresa": name,
        "documentos": documents,
        "erros": errors,
        "mensagem": message,
        "nsu_final": 0,
    }


def _export_name_for_company(company: dict[str, Any], save_mode_slug: str) -> str:
    if "cnpj" in save_mode_slug:
        cnpj = "".join(char for char in str(company.get("cnpj", "") or "") if char.isalnum())
        if cnpj:
            return f"{cnpj}.zip"
    return f"{company['cod']}.zip"


def _is_certificate_expired(company: dict[str, Any]) -> bool:
    venc = str(company.get("venc", "") or "").strip()
    if not venc:
        return False

    try:
        expiry = time.strptime(venc, "%d/%m/%Y")
    except ValueError:
        return False

    return time.localtime() > expiry


def _extract_nsu_from_message(message: str) -> int | None:
    for marker in ("Consultando NSU:", "NSU:"):
        if marker not in message:
            continue

        try:
            raw = message.split(marker, 1)[1].strip().split()[0]
            return int(raw)
        except (IndexError, ValueError):
            return None

    return None


def _slug_text(value: Any) -> str:
    text = str(value or "").strip().lower()
    ascii_text = normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return ascii_text


def _emit(callback: EventCallback | None, **payload: Any) -> None:
    if callback is None:
        return
    callback(payload)
