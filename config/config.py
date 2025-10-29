from __future__ import annotations
import json
import logging
import os
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

# Status de parada
STATUS_STOP = [204, 400]
MAX_TENT = 2

## ------------------------------------------------------------------------------
## Diretórios padrão
## ------------------------------------------------------------------------------
def get_base_dir() -> Path:
    """Retorna o diretório base da aplicação."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent

ROOT_DIR = get_base_dir()

# Configuração de diretórios usando Path para melhor manipulação
_DIR_PATHS = {
    'certificados': ROOT_DIR / 'cert_path',
    'notas': ROOT_DIR / 'packs',
    'logs': ROOT_DIR / 'logs',
    'config': ROOT_DIR / 'config',
    'docs': ROOT_DIR / 'docs',
    'EVENTOS':  ROOT_DIR / 'packs' / 'EVENTOS',
    'TOMADOS':  ROOT_DIR / 'packs' / 'TOMADOS',
    'PRESTADOS':  ROOT_DIR / 'packs' / 'PRESTADOS'
}

_DIR_FILES = {
    'planilha_modelo': _DIR_PATHS['config'] / 'relatorio.xlsm',
    'config_json': _DIR_PATHS['config'] / 'config.json',
    'cadastros_json': _DIR_PATHS['config'] / 'cadastros.json',
    'icone': _DIR_PATHS['config'] / 'icone.ico',
    'instrucoes': _DIR_PATHS['docs'] / 'index.md',
    'versao': _DIR_PATHS['docs'] / 'versao.txt',
}

# Combinar todos os diretórios em um único dicionário
DIRETORIOS = {**_DIR_PATHS, **_DIR_FILES}

## ------------------------------------------------------------------------------
## Carregamento de configuração padrão
## ------------------------------------------------------------------------------
@dataclass
class Config:
    file_prefix: str = "NFS-e"
    download_pdf: bool = False
    delay_seconds: float = 0.5
    timeout: int = 60

    @classmethod
    def load(cls, path: str | Path) -> Config:
        """Carrega configuração de ``path`` ou cria com valores padrão."""
        path = Path(path)
        
        if path.exists():
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        else:
            data = {}
            # Cria arquivo com configurações padrão
            default_config = cls()
            default_config.save(path)
            return default_config

        # Combina configurações do arquivo com valores padrão
        cfg_data = asdict(cls())
        cfg_data.update(data)
        return cls(**cfg_data)

    def save(self, path: str | Path) -> None:
        """Salva configuração em ``path`` como JSON."""
        path = Path(path)
        
        with path.open("w", encoding="utf-8") as f:
            json.dump(asdict(self), f, indent=2, ensure_ascii=False)

## ------------------------------------------------------------------------------
## Configurações de Logging
## ------------------------------------------------------------------------------
class LogConfig:
    """Classe para configuração centralizada de logging."""
    
    _CONFIGURADO = False
    
    @classmethod
    def configurar(cls, nome_arquivo: str | None = None, nivel: int = logging.INFO) -> str:
        """
        Configura o sistema de logging para o projeto.
        
        Args:
            nome_arquivo: Nome do arquivo de log (opcional)
            nivel: Nível de logging (INFO, DEBUG, etc.)
            
        Returns:
            Caminho completo do arquivo de log
        """
        if cls._CONFIGURADO:
            return cls._caminho_log  # type: ignore
        
        # Nome do arquivo de log com timestamp
        if nome_arquivo is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"nfse_downloader_{timestamp}.log"
        
        cls._caminho_log = DIRETORIOS['logs'] / nome_arquivo
        
        # Configurar formato
        formato = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - [%(threadName)s] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Obter o logger raiz
        logger = logging.getLogger()
        logger.setLevel(nivel)
        
        # Remover handlers existentes para evitar duplicação
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # Handler para arquivo
        file_handler = logging.FileHandler(cls._caminho_log, encoding='utf-8', mode='a')
        file_handler.setFormatter(formato)
        file_handler.setLevel(nivel)
        
        # Handler para terminal
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formato)
        console_handler.setLevel(nivel)
        
        # Adicionar handlers
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        # Log inicial
        logging.info("=== SISTEMA DE LOG INICIADO ===")
        logging.info("Arquivo de log: %s", cls._caminho_log)
        logging.info("Nível de log: %s", logging.getLevelName(nivel))
        
        cls._CONFIGURADO = True
        return str(cls._caminho_log)

# Alias para manter compatibilidade
configurar_logging = LogConfig.configurar

def obter_logger(nome: str) -> logging.Logger:
    """
    Função auxiliar para obter um logger com nome específico.
    Útil para módulos que precisam de seu próprio logger.
    """
    return logging.getLogger(nome)
