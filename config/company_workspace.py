from __future__ import annotations

import logging
import re
import shutil
from pathlib import Path

from config.config import DIRETORIOS

logger = logging.getLogger(__name__)

INVALID_FS_CHARS = r'[<>:"/\\|?*]'


def sanitize_company_folder_name(company_name: str | None, cod_empresa: str | int) -> str:
    text = str(company_name or "").strip()
    if text:
        text = re.sub(INVALID_FS_CHARS, " ", text)
        text = " ".join(text.split()).strip(" .")
    return text or f"Empresa {cod_empresa}"


def company_root_dir(cod_empresa: str | int, company_name: str | None) -> Path:
    folder_name = sanitize_company_folder_name(company_name, cod_empresa)
    return Path(DIRETORIOS["notas"]) / folder_name


def company_period_dir(cod_empresa: str | int, company_name: str | None, year: str | int, month: str | int) -> Path:
    return company_root_dir(cod_empresa, company_name) / competence_folder_name(year, month)


def company_control_file(cod_empresa: str | int, company_name: str | None) -> Path:
    return company_root_dir(cod_empresa, company_name) / "nsu_competencia.json"


def competence_folder_name(year: str | int, month: str | int) -> str:
    return f"{int(month):02d}-{int(year):04d}"


def ensure_company_workspace(
    cod_empresa: str | int,
    cnpj: str,
    company_name: str | None,
    previous_company_name: str | None = None,
) -> tuple[bool, str]:
    cod_str = str(cod_empresa)
    template_dir = Path(DIRETORIOS["notas"]) / "0"
    target_dir = company_root_dir(cod_empresa, company_name)
    legacy_code_dir = Path(DIRETORIOS["notas"]) / cod_str
    previous_name_dir = (
        company_root_dir(cod_empresa, previous_company_name)
        if previous_company_name and previous_company_name != company_name
        else None
    )

    try:
        migration_source = None
        if target_dir.exists():
            migration_source = None
        elif previous_name_dir and previous_name_dir.exists():
            migration_source = previous_name_dir
        elif legacy_code_dir.exists() and legacy_code_dir != target_dir:
            migration_source = legacy_code_dir

        if migration_source is not None and migration_source.exists():
            shutil.move(str(migration_source), str(target_dir))
            logger.info("Pasta da empresa %s migrada de %s para %s.", cod_str, migration_source.name, target_dir.name)

        if not target_dir.exists():
            if template_dir.exists() and any(template_dir.iterdir()):
                shutil.copytree(template_dir, target_dir)
                logger.info("Estrutura da empresa %s criada a partir da pasta modelo.", cod_str)
            else:
                target_dir.mkdir(parents=True, exist_ok=True)
                logger.info("Pasta raiz da empresa %s criada sem modelo base.", cod_str)

        _remove_legacy_spreadsheets(target_dir)
        return True, str(target_dir)

    except Exception as exc:
        logger.error("Erro ao preparar estrutura da empresa %s: %s", cod_str, exc)
        return False, f"Erro ao preparar estrutura da empresa {cod_str}: {exc}"


def ensure_period_workspace(
    cod_empresa: str | int,
    cnpj: str,
    company_name: str | None,
    year: str | int,
    month: str | int,
) -> tuple[bool, str]:
    ok, message = ensure_company_workspace(cod_empresa, cnpj, company_name)
    if not ok:
        return False, message

    root_dir = Path(message)
    period_dir = root_dir / competence_folder_name(year, month)

    try:
        period_dir.mkdir(parents=True, exist_ok=True)
        for folder_name in ("PRESTADOS", "TOMADOS", "EVENTOS"):
            folder = period_dir / folder_name
            folder.mkdir(parents=True, exist_ok=True)
            if folder_name in {"PRESTADOS", "TOMADOS"}:
                (folder / "Canceladas").mkdir(parents=True, exist_ok=True)

        _remove_legacy_spreadsheets(period_dir)
        return True, str(period_dir)
    except Exception as exc:
        logger.error("Erro ao preparar pasta da competencia %s/%s para a empresa %s: %s", month, year, cod_empresa, exc)
        return False, f"Erro ao preparar pasta da competencia {month}-{year} da empresa {cod_empresa}: {exc}"


def _remove_legacy_spreadsheets(base_dir: Path) -> None:
    for file_path in base_dir.glob("*.xlsm"):
        try:
            file_path.unlink()
            logger.info("Planilha legada removida de %s.", file_path)
        except OSError as exc:
            logger.warning("Nao foi possivel remover a planilha legada %s: %s", file_path, exc)
