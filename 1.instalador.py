import os
import sys
import subprocess
import platform
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

def get_base_dir() -> Path:
    """Retorna o diretório base da aplicação."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

ROOT_DIR = get_base_dir()

# Configuração de diretórios usando Path para melhor manipulação
_DIR_PATHS = {
    'certificados': ROOT_DIR / 'cert_path',
    'notas': ROOT_DIR / 'packs',
    'logs': ROOT_DIR / 'logs',
    'config': ROOT_DIR / 'config',
    'docs': ROOT_DIR / 'docs',
    'EVENTOS':  ROOT_DIR / 'packs' / '0' / 'EVENTOS',
    'TOMADOS':  ROOT_DIR / 'packs' / '0' / 'TOMADOS',
    'PRESTADOS':  ROOT_DIR / 'packs' / '0' / 'PRESTADOS'
}

_DIR_FILES = {
    'packs': ROOT_DIR / '_internal' / 'packs',
    'docs': ROOT_DIR / '_internal' / 'docs',

}

_CONFIG_FILES = {
    'cadastros_json': ROOT_DIR / '_internal' / 'config' / 'cadastros.json',
    'icone': ROOT_DIR / '_internal' / 'config' / 'icone.ico',
    'config_json': ROOT_DIR / '_internal' / 'config' / 'config.json' 
}

# Criar diretórios necessários ao importar o módulo
def _inicializar_diretorios() -> None:
    """Cria os diretórios necessários para a aplicação."""
    for diretorio in _DIR_PATHS.values():
        diretorio.mkdir(parents=True, exist_ok=True)

def _realocar_root_dir() -> None:
    """Move diretórios e arquivos para os novos locais"""
    # Mover _DIR_FILES para ROOT_DIR
    for key, source_path in _DIR_FILES.items():
        dest_path = ROOT_DIR / key
        try:
            if source_path.exists():
                if dest_path.exists():
                    shutil.rmtree(dest_path)
                shutil.move(str(source_path), str(dest_path))
                print(f"✓ Movido {source_path} -> {dest_path}")
        except Exception as e:
            print(f"✗ Erro ao mover {source_path}: {str(e)}")

    # Mover _CONFIG_FILES para _DIR_PATHS['config']
    config_dir = _DIR_PATHS['config']
    for key, source_path in _CONFIG_FILES.items():
        dest_path = config_dir / source_path.name
        try:
            if source_path.exists():
                if dest_path.exists():
                    dest_path.unlink()
                shutil.move(str(source_path), str(dest_path))
                print(f"✓ Movido {source_path} -> {dest_path}")
        except Exception as e:
            print(f"✗ Erro ao mover {source_path}: {str(e)}")

    # Remover diretório _internal se estiver vazio
    internal_dir = ROOT_DIR / '_internal'
    try:
        if internal_dir.exists() and not any(internal_dir.iterdir()):
            internal_dir.rmdir()
            print("✓ Diretório _internal removido")
    except Exception as e:
        print(f"✗ Erro ao remover _internal: {str(e)}")

def atualizar_pip():
    """Atualiza o pip antes de outras instalações"""
    print("\nAtualizando pip...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("✓ Pip atualizado com sucesso")
    except subprocess.CalledProcessError as e:
        print(f"✗ Falha ao atualizar pip: {str(e)}")

def carregar_requirements():
    """Carrega as dependências do arquivo requirements.txt"""
    requirements_path = os.path.join(ROOT_DIR, 'requirements.txt')
    
    if not os.path.exists(requirements_path):
        print(f"✗ Arquivo requirements.txt não encontrado em: {requirements_path}")
        return None
    
    try:
        with open(requirements_path, 'r', encoding='utf-8') as file:
            dependencias = []
            for line in file:
                line = line.strip()
                # Ignora linhas vazias, comentários e opções de instalação
                if line and not line.startswith('#') and not line.startswith('-'):
                    dependencias.append(line)
            return dependencias
    except Exception as e:
        print(f"✗ Erro ao ler requirements.txt: {str(e)}")
        return None

def verificar_instalar_dependencias():
    """Verifica e instala todas as dependências necessárias"""
    dependencias = carregar_requirements()
    
    if dependencias is None:
        return False, ["Falha ao carregar requirements.txt"]

    sistema_operacional = platform.system()
    python_exec = sys.executable
    falhas = []
    
    print("\n=== VERIFICANDO DEPENDÊNCIAS ===")
    print(f"Sistema Operacional: {sistema_operacional}")
    print(f"Python: {python_exec}")
    print(f"Encontradas {len(dependencias)} dependências no requirements.txt")
    
    try:
        import pip
        atualizar_pip()
    except ImportError:
        print("\nERRO: Pip não está instalado. Instale o pip primeiro.")
        return False
    
    # Verifica e instala cada dependência
    for pacote in dependencias:
        # Extrai o nome base do pacote (remove versão e extras)
        nome_base = pacote.split('==')[0].split('>=')[0].split('<=')[0].split('[')[0].strip()
        
        try:
            __import__(nome_base)
            print(f"✔ {pacote} já está instalado")
        except ImportError:
            print(f"\nInstalando {pacote}...")
            try:
                subprocess.check_call([python_exec, "-m", "pip", "install", pacote], stdout=subprocess.DEVNULL)
                print(f"✔ {pacote} instalado com sucesso")
            except subprocess.CalledProcessError:
                print(f"✖ Falha ao instalar {pacote}")
                falhas.append(pacote)
    
    return len(falhas) == 0, falhas

def formatar_lista_falhas(falhas):
    """Formata a lista de falhas com bullets e quebras de linha"""
    if not falhas:
        return "Nenhuma falha encontrada"
    
    lista_formatada = ""
    for i, falha in enumerate(falhas, 1):
        lista_formatada += f"• {falha}"
        if i < len(falhas):
            lista_formatada += "\n"
    
    return lista_formatada
            
def mostrar_popup(mensagem, titulo="Instalação"):
    """Mostra popup informativo usando Tkinter"""
    try:
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal
        messagebox.showinfo(titulo, mensagem)
        root.destroy()
    except Exception as e:
        print(f"\nErro ao exibir popup: {str(e)}")
        print(f"\n=== {titulo} ===\n{mensagem}\n")

def main():
    print("=== CONFIGURADOR AUTOMÁTICO ===")
    print("Criando estrutura de pastas...")
    _inicializar_diretorios()
    _realocar_root_dir()
    
    sucesso, falhas = verificar_instalar_dependencias()
    
    if sucesso:
        mensagem = "Todas as dependências foram instaladas com sucesso!\n\nAgora você pode executar o programa principal."
        print(f"\n{mensagem}")
        mostrar_popup(mensagem, "Instalação Completa")
    else:
        lista_falhas = formatar_lista_falhas(falhas)
        mensagem = (
            "Houve problemas na instalação das dependências.\n\n"
            "Dependências com falha:\n"
            f"{lista_falhas}\n\n"
            "Consulte as mensagens acima para corrigir."
        )
        
        # Versão para console
        print("\n" + "="*50)
        print("ERROS ENCONTRADOS:")
        print(lista_falhas.replace("• ", "- "))
        print("="*50)
        
        # Versão para popup
        mostrar_popup(mensagem, "Erro na Instalação")
        sys.exit(1)

if __name__ == "__main__":
    main()