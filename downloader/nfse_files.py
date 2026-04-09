from __future__ import annotations

import re
import shutil
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path


INVALID_FS_CHARS = r'[<>:"/\\|?*]'


@dataclass(slots=True)
class NFSeFileInfo:
    numero: str
    chave: str
    retido: bool

    @property
    def basename(self) -> str:
        suffix = " - Retido" if self.retido else ""
        return sanitize_file_stem(f"{self.numero} - {self.chave}{suffix}")


def build_nfse_file_info(xml_bytes: bytes, chave: str, fallback_number: str | int) -> NFSeFileInfo:
    numero = str(fallback_number)
    retido = False

    try:
        root = ET.fromstring(xml_bytes)

        numero_node = root.find(".//{*}nNFSe")
        if numero_node is not None and (numero_node.text or "").strip():
            numero = (numero_node.text or "").strip()

        for element in root.iter():
            local_name = _local_name(element.tag)
            if local_name != "vTotalRet":
                continue
            if _parse_decimal(element.text) > 0:
                retido = True
                break
    except ET.ParseError:
        pass

    numero_digits = "".join(char for char in str(numero) if char.isdigit())
    numero_final = numero_digits or sanitize_file_stem(str(numero)) or str(fallback_number)
    return NFSeFileInfo(numero=numero_final, chave=str(chave), retido=retido)


def move_document_to_canceladas(file_path: str | Path, canceladas_dir: str | Path) -> None:
    source = Path(file_path)
    if not source.exists():
        return

    target_dir = Path(canceladas_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / source.name

    if target.exists():
        target.unlink()

    shutil.move(str(source), str(target))


def sanitize_file_stem(value: str) -> str:
    text = re.sub(INVALID_FS_CHARS, " ", str(value or ""))
    text = " ".join(text.split()).strip(" .")
    return text or "Documento"


def _local_name(tag: str) -> str:
    if "}" in str(tag):
        return str(tag).rsplit("}", 1)[-1]
    return str(tag)


def _parse_decimal(value: str | None) -> float:
    raw = str(value or "").strip()
    if not raw:
        return 0.0

    if "," in raw and "." in raw:
        normalized = raw.replace(".", "").replace(",", ".")
    elif "," in raw:
        normalized = raw.replace(",", ".")
    else:
        normalized = raw

    try:
        return float(normalized)
    except ValueError:
        return 0.0
