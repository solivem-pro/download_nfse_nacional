from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from config.cadastro_db import list_companies
from config.config import Config, DIRETORIOS
from config.utils import formatar_cnpj

CONSULT_MODE_COMPETENCIA = "Compet\u00eancia"
CONSULT_MODE_EMISSAO = "Emiss\u00e3o"
SAVE_MODE_CODIGO = "C\u00f3digo"
SAVE_MODE_CNPJ = "CNPJ"


def load_app_config() -> Config:
    return Config.load(DIRETORIOS["config_json"])


def save_app_config(data: dict[str, Any]) -> Config:
    normalized = dict(data)
    normalized["consult_mode"] = _normalize_consult_mode(normalized.get("consult_mode"))
    normalized["save_mode"] = _normalize_save_mode(normalized.get("save_mode"))
    config = Config(**normalized)
    config.save(DIRETORIOS["config_json"])
    return config


def config_to_dict(config: Config) -> dict[str, Any]:
    data = asdict(config)
    data["consult_mode"] = _normalize_consult_mode(data.get("consult_mode"))
    data["save_mode"] = _normalize_save_mode(data.get("save_mode"))
    return data


def load_company_rows() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for company in list_companies():
        raw_cnpj = str(company.get("cnpj", "") or "")
        venc = str(company.get("venc", "") or "")
        expired, warning = _certificate_state(venc)

        rows.append(
            {
                "cod": company.get("cod", ""),
                "empresa": str(company.get("empresa", "") or "").strip(),
                "cnpj": raw_cnpj,
                "cnpj_formatado": formatar_cnpj(raw_cnpj),
                "cert_path": str(company.get("cert_path", "") or "").strip(),
                "cert_pass": str(company.get("cert_pass", "") or "").strip(),
                "venc": venc,
                "cert_expired": expired,
                "cert_warning": warning,
                "cert_status": _certificate_status_label(expired, warning, venc),
            }
        )

    rows.sort(key=_company_sort_key)
    return rows


def summarize_company_rows(rows: list[dict[str, Any]]) -> dict[str, int]:
    total = len(rows)
    expired = sum(1 for row in rows if row.get("cert_expired"))
    warning = sum(1 for row in rows if row.get("cert_warning"))
    valid = max(0, total - expired - warning)

    return {
        "total": total,
        "expired": expired,
        "warning": warning,
        "valid": valid,
    }


def available_years() -> list[str]:
    current = datetime.now().year
    return [str(year) for year in range(current - 5, current + 1)]


def default_period() -> tuple[str, str]:
    today = datetime.now()
    last_month = today.replace(day=1) - timedelta(days=1)
    return str(last_month.year), f"{last_month.month:02d}"


def read_markdown_document() -> str:
    doc_path = Path(DIRETORIOS["instrucoes"])
    if not doc_path.exists():
        return "Documenta\u00e7\u00e3o n\u00e3o encontrada."
    return doc_path.read_text(encoding="utf-8")


def app_version() -> str:
    try:
        from docs.version import __version__

        return __version__
    except Exception:
        return "0.0.0"


def _certificate_state(vencimento: str) -> tuple[bool, bool]:
    vencimento = (vencimento or "").strip()
    if not vencimento:
        return False, False

    try:
        target = datetime.strptime(vencimento, "%d/%m/%Y")
    except ValueError:
        return False, False

    now = datetime.now()
    delta_days = (target - now).days
    if delta_days < 0:
        return True, False
    if delta_days <= 30:
        return False, True
    return False, False


def _certificate_status_label(expired: bool, warning: bool, vencimento: str) -> str:
    if expired:
        return f"Vencido ({vencimento})"
    if warning:
        return f"Vence em breve ({vencimento})"
    if vencimento:
        return f"V\u00e1lido at\u00e9 {vencimento}"
    return "Sem vencimento informado"


def _normalize_consult_mode(value: Any) -> str:
    text = str(value or "").strip()
    if text in {CONSULT_MODE_COMPETENCIA, "Competencia", "CompetÃªncia", "CompetÃƒÂªncia"}:
        return CONSULT_MODE_COMPETENCIA
    if text in {CONSULT_MODE_EMISSAO, "Emissao", "EmissÃ£o", "EmissÃƒÂ£o"}:
        return CONSULT_MODE_EMISSAO
    return CONSULT_MODE_COMPETENCIA


def _normalize_save_mode(value: Any) -> str:
    text = str(value or "").strip()
    if text in {SAVE_MODE_CODIGO, "Codigo", "CÃ³digo", "CÃƒÂ³digo"}:
        return SAVE_MODE_CODIGO
    if text.upper() == SAVE_MODE_CNPJ:
        return SAVE_MODE_CNPJ
    return SAVE_MODE_CODIGO


def _company_sort_key(row: dict[str, Any]) -> tuple[int, str]:
    raw = str(row.get("cod", "") or "").strip()
    if raw.isdigit():
        return (0, f"{int(raw):010d}")
    return (1, raw.lower())
