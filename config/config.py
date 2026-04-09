from __future__ import annotations

import json
import logging
import os
import shutil
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

# Status de parada
STATUS_STOP = [204, 400]
MAX_TENT = 2
APP_RUNTIME_NAME = "Portal NFSe"
APP_DOWNLOADS_NAME = "Portal NFSe"


def get_resource_dir() -> Path:
    """Return the directory that contains bundled resources."""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _desktop_dir() -> Path:
    desktop_dir = Path.home() / "Desktop"
    if desktop_dir.exists():
        return desktop_dir
    return Path.home()


def _appdata_dir() -> Path:
    appdata = os.getenv("APPDATA", "").strip()
    if appdata:
        return Path(appdata)
    return Path.home() / "AppData" / "Roaming"


def _looks_like_source_checkout(path: Path) -> bool:
    if not path.exists():
        return False

    markers = (
        path / ".git",
        path / "download_nfse_qt.py",
        path / "requirements.txt",
        path / "_clean_export",
    )
    return any(marker.exists() for marker in markers)


def get_base_dir() -> Path:
    """Return the internal runtime directory used by the packaged app."""
    if getattr(sys, "frozen", False):
        return _appdata_dir() / APP_RUNTIME_NAME
    return _repo_root()


def get_downloads_dir() -> Path:
    """
    Return the user-visible downloads directory.

    In the packaged app, downloads stay on the Desktop while internal files go
    to AppData. During development we keep the existing repository layout.
    """
    if getattr(sys, "frozen", False):
        preferred = _desktop_dir() / APP_DOWNLOADS_NAME
        if _looks_like_source_checkout(preferred):
            return _desktop_dir() / f"{APP_DOWNLOADS_NAME} Downloads"
        return preferred
    return _repo_root() / "packs"


RESOURCE_DIR = get_resource_dir()
ROOT_DIR = get_base_dir()
DOWNLOADS_DIR = get_downloads_dir()

_DIR_PATHS = {
    "certificados": ROOT_DIR / "cert_path",
    "notas": DOWNLOADS_DIR,
    "logs": ROOT_DIR / "logs",
    "config": ROOT_DIR / "config",
    "docs": ROOT_DIR / "docs",
    "packs": DOWNLOADS_DIR,
    "EVENTOS": DOWNLOADS_DIR / "EVENTOS",
    "TOMADOS": DOWNLOADS_DIR / "TOMADOS",
    "PRESTADOS": DOWNLOADS_DIR / "PRESTADOS",
    "temp": ROOT_DIR / "temp",
}

_DIR_FILES = {
    "cadastros_db": _DIR_PATHS["config"] / "cadastros.db",
    "config_json": _DIR_PATHS["config"] / "config.json",
    "cadastros_json": _DIR_PATHS["config"] / "cadastros.json",
    "icone": _DIR_PATHS["config"] / "icone.ico",
    "instrucoes": ROOT_DIR / "README.md",
    "versao": _DIR_PATHS["docs"] / "versao.txt",
    "vba": _DIR_PATHS["docs"] / "vba.bas",
}

DIRETORIOS = {**_DIR_PATHS, **_DIR_FILES}


def _copy_file_if_missing(source: Path, target: Path) -> None:
    if not source.exists() or target.exists():
        return

    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)


def _copy_tree_if_missing(source_dir: Path, target_dir: Path) -> None:
    if not source_dir.exists():
        return

    target_dir.mkdir(parents=True, exist_ok=True)

    for source_path in source_dir.rglob("*"):
        relative_path = source_path.relative_to(source_dir)
        target_path = target_dir / relative_path

        if source_path.is_dir():
            target_path.mkdir(parents=True, exist_ok=True)
            continue

        if target_path.exists():
            continue

        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, target_path)


def _migrate_legacy_desktop_workspace() -> None:
    """
    Migrate data from the previous desktop-based layout to the new AppData
    layout without overwriting user data.
    """
    if not getattr(sys, "frozen", False):
        return

    legacy_root = _desktop_dir() / APP_DOWNLOADS_NAME
    if not legacy_root.exists():
        return

    _copy_tree_if_missing(legacy_root / "cert_path", ROOT_DIR / "cert_path")
    _copy_tree_if_missing(legacy_root / "logs", ROOT_DIR / "logs")
    _copy_tree_if_missing(legacy_root / "temp", ROOT_DIR / "temp")

    for relative_file in (
        Path("config") / "cadastros.db",
        Path("config") / "cadastros.json",
        Path("config") / "config.json",
        Path("config") / "icone.ico",
        Path("docs") / "qt_primeira_execucao.md",
        Path("docs") / "requirements.txt",
        Path("docs") / "vba.bas",
        Path("docs") / "version.py",
        Path("docs") / "version_file.txt",
        Path("README.md"),
    ):
        _copy_file_if_missing(legacy_root / relative_file, ROOT_DIR / relative_file)

    _copy_tree_if_missing(legacy_root / "packs", DOWNLOADS_DIR)


def initialize_runtime_environment() -> None:
    """
    Prepare the runtime workspace.

    A one-file build extracts its bundled resources to a temp directory. On the
    first run we seed the user's Desktop workspace with the minimal set of files
    and folders required by the app, without overwriting existing data.
    """
    for key in ("certificados", "notas", "logs", "config", "docs", "temp"):
        Path(DIRETORIOS[key]).mkdir(parents=True, exist_ok=True)

    _migrate_legacy_desktop_workspace()
    _copy_tree_if_missing(RESOURCE_DIR / "cert_path", Path(DIRETORIOS["certificados"]))
    _copy_tree_if_missing(RESOURCE_DIR / "packs", Path(DIRETORIOS["packs"]))

    for relative_file in (
        Path("config") / "cadastros.db",
        Path("config") / "cadastros.json",
        Path("config") / "config.json",
        Path("config") / "icone.ico",
        Path("docs") / "qt_primeira_execucao.md",
        Path("docs") / "requirements.txt",
        Path("docs") / "vba.bas",
        Path("docs") / "version.py",
        Path("docs") / "version_file.txt",
        Path("README.md"),
    ):
        _copy_file_if_missing(RESOURCE_DIR / relative_file, ROOT_DIR / relative_file)


@dataclass
class Config:
    file_prefix: str = "NSU"
    download_pdf: bool = False
    delay_seconds: float = 0.5
    timeout: int = 60
    consult_mode: str = "Competencia"
    save_mode: str = "Codigo"

    @classmethod
    def load(cls, path: str | Path) -> "Config":
        """Load config from path or create a default config file."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        if path.exists():
            with path.open("r", encoding="utf-8") as file:
                data = json.load(file)
        else:
            data = {}
            default_config = cls()
            default_config.save(path)
            return default_config

        cfg_data = asdict(cls())
        cfg_data.update(data)
        return cls(**cfg_data)

    def save(self, path: str | Path) -> None:
        """Persist config as JSON."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        with path.open("w", encoding="utf-8") as file:
            json.dump(asdict(self), file, indent=2, ensure_ascii=False)


class LogConfig:
    """Central logging setup for the project."""

    _CONFIGURADO = False
    _caminho_log: Path

    @classmethod
    def configurar(cls, nome_arquivo: str | None = None, nivel: int = logging.INFO) -> str:
        if cls._CONFIGURADO:
            return str(cls._caminho_log)

        if nome_arquivo is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"nfse_downloader_{timestamp}.log"

        Path(DIRETORIOS["logs"]).mkdir(parents=True, exist_ok=True)
        cls._caminho_log = Path(DIRETORIOS["logs"]) / nome_arquivo

        formato = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - [%(threadName)s] - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

        logger = logging.getLogger()
        logger.setLevel(nivel)

        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        file_handler = logging.FileHandler(cls._caminho_log, encoding="utf-8", mode="a")
        file_handler.setFormatter(formato)
        file_handler.setLevel(nivel)

        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formato)
        console_handler.setLevel(nivel)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        logging.info("=== SISTEMA DE LOG INICIADO ===")
        logging.info("Arquivo de log: %s", cls._caminho_log)
        logging.info("Nivel de log: %s", logging.getLevelName(nivel))

        cls._CONFIGURADO = True
        return str(cls._caminho_log)


configurar_logging = LogConfig.configurar


def obter_logger(nome: str) -> logging.Logger:
    return logging.getLogger(nome)
