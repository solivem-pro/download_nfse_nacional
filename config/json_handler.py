import json
import os
from config.config import DIRETORIOS

def carregar_json(caminho):
    """Carrega dados de um arquivo JSON"""
    if not os.path.exists(caminho):
        return {}
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_json(dados, caminho):
    """Salva dados em um arquivo JSON"""
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)

def carregar_cadastros():
    return carregar_json(DIRETORIOS['cadastros_json'])

def salvar_cadastros(dados):
    salvar_json(dados, DIRETORIOS['cadastros_json'])