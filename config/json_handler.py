from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from config.cadastro_db import load_cadastros_payload, save_cadastros_payload
from config.config import DIRETORIOS


def _is_cadastros_path(caminho: str | os.PathLike[str]) -> bool:
    try:
        return Path(caminho).resolve() == Path(DIRETORIOS["cadastros_json"]).resolve()
    except OSError:
        return os.fspath(caminho) == os.fspath(DIRETORIOS["cadastros_json"])


def carregar_json(caminho: str | os.PathLike[str]) -> dict[str, Any]:
    """Carrega dados de um arquivo JSON."""
    if _is_cadastros_path(caminho):
        return load_cadastros_payload()

    if not os.path.exists(caminho):
        return {}

    with open(caminho, "r", encoding="utf-8") as file:
        return json.load(file)


def salvar_json(dados: dict[str, Any], caminho: str | os.PathLike[str]) -> None:
    """Salva dados em um arquivo JSON."""
    if _is_cadastros_path(caminho):
        save_cadastros_payload(dados)
        return

    os.makedirs(os.path.dirname(os.fspath(caminho)), exist_ok=True)
    with open(caminho, "w", encoding="utf-8") as file:
        json.dump(dados, file, indent=2, ensure_ascii=False)


def carregar_cadastros() -> dict[str, Any]:
    return load_cadastros_payload()


def salvar_cadastros(dados: dict[str, Any]) -> None:
    save_cadastros_payload(dados)
