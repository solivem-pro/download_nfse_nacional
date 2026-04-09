import ctypes
import platform
import shutil
import subprocess
import sys
from pathlib import Path


def get_base_dir() -> Path:
    """Retorna o diretorio base da aplicacao."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


ROOT_DIR = get_base_dir()

_DIR_PATHS = {
    "certificados": ROOT_DIR / "cert_path",
    "notas": ROOT_DIR / "packs",
    "logs": ROOT_DIR / "logs",
    "config": ROOT_DIR / "config",
    "docs": ROOT_DIR / "docs",
    "EVENTOS": ROOT_DIR / "packs" / "0" / "EVENTOS",
    "TOMADOS": ROOT_DIR / "packs" / "0" / "TOMADOS",
    "PRESTADOS": ROOT_DIR / "packs" / "0" / "PRESTADOS",
    "temp": ROOT_DIR / "temp",
}

_DIR_FILES = {
    "packs": ROOT_DIR / "_internal" / "packs",
    "docs": ROOT_DIR / "_internal" / "docs",
}

_CONFIG_FILES = {
    "cadastros_db": ROOT_DIR / "_internal" / "config" / "cadastros.db",
    "cadastros_json": ROOT_DIR / "_internal" / "config" / "cadastros.json",
    "icone": ROOT_DIR / "_internal" / "config" / "icone.ico",
    "config_json": ROOT_DIR / "_internal" / "config" / "config.json",
}

IMPORT_NAME_OVERRIDES = {
    "python-dateutil": "dateutil",
}


def _inicializar_diretorios() -> None:
    """Cria os diretorios necessarios para a aplicacao."""
    for diretorio in _DIR_PATHS.values():
        diretorio.mkdir(parents=True, exist_ok=True)


def _realocar_root_dir() -> None:
    """Move diretorios e arquivos internos para os locais finais."""
    for source_path in _DIR_FILES.values():
        dest_path = ROOT_DIR / source_path.name
        try:
            if source_path.exists():
                if dest_path.exists():
                    shutil.rmtree(dest_path)
                shutil.move(str(source_path), str(dest_path))
                print(f"[OK] Movido {source_path} -> {dest_path}")
        except Exception as exc:
            print(f"[ERRO] Falha ao mover {source_path}: {exc}")

    config_dir = _DIR_PATHS["config"]
    for source_path in _CONFIG_FILES.values():
        dest_path = config_dir / source_path.name
        try:
            if source_path.exists():
                if dest_path.exists():
                    dest_path.unlink()
                shutil.move(str(source_path), str(dest_path))
                print(f"[OK] Movido {source_path} -> {dest_path}")
        except Exception as exc:
            print(f"[ERRO] Falha ao mover {source_path}: {exc}")

    internal_dir = ROOT_DIR / "_internal"
    try:
        if internal_dir.exists() and not any(internal_dir.iterdir()):
            internal_dir.rmdir()
            print("[OK] Diretorio _internal removido")
    except Exception as exc:
        print(f"[ERRO] Falha ao remover _internal: {exc}")


def atualizar_pip() -> None:
    """Atualiza o pip antes das demais instalacoes."""
    print("\nAtualizando pip...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("[OK] Pip atualizado com sucesso")
    except subprocess.CalledProcessError as exc:
        print(f"[ERRO] Falha ao atualizar pip: {exc}")


def carregar_requirements():
    """Carrega as dependencias do arquivo requirements.txt."""
    requirements_path = ROOT_DIR / "requirements.txt"
    if not requirements_path.exists():
        print(f"[ERRO] Arquivo requirements.txt nao encontrado em: {requirements_path}")
        return None

    try:
        dependencias = []
        with requirements_path.open("r", encoding="utf-8") as file:
            for line in file:
                line = line.strip()
                if line and not line.startswith("#") and not line.startswith("-"):
                    dependencias.append(line)
        return dependencias
    except Exception as exc:
        print(f"[ERRO] Erro ao ler requirements.txt: {exc}")
        return None


def verificar_instalar_dependencias():
    """Verifica e instala todas as dependencias necessarias."""
    dependencias = carregar_requirements()
    if dependencias is None:
        return False, ["Falha ao carregar requirements.txt"]

    sistema_operacional = platform.system()
    python_exec = sys.executable
    falhas = []

    print("\n=== VERIFICANDO DEPENDENCIAS ===")
    print(f"Sistema Operacional: {sistema_operacional}")
    print(f"Python: {python_exec}")
    print(f"Encontradas {len(dependencias)} dependencias no requirements.txt")

    try:
        import pip  # noqa: F401

        atualizar_pip()
    except ImportError:
        print("\nERRO: Pip nao esta instalado. Instale o pip primeiro.")
        return False, ["pip nao instalado"]

    for pacote in dependencias:
        nome_base = pacote.split("==")[0].split(">=")[0].split("<=")[0].split("[")[0].strip()
        modulo_teste = IMPORT_NAME_OVERRIDES.get(nome_base, nome_base.replace("-", "_"))
        try:
            __import__(modulo_teste)
            print(f"[OK] {pacote} ja esta instalado")
        except ImportError:
            print(f"\nInstalando {pacote}...")
            try:
                subprocess.check_call([python_exec, "-m", "pip", "install", pacote], stdout=subprocess.DEVNULL)
                print(f"[OK] {pacote} instalado com sucesso")
            except subprocess.CalledProcessError:
                print(f"[ERRO] Falha ao instalar {pacote}")
                falhas.append(pacote)

    return len(falhas) == 0, falhas


def formatar_lista_falhas(falhas):
    """Formata a lista de falhas com bullets simples."""
    if not falhas:
        return "Nenhuma falha encontrada"
    return "\n".join(f"- {falha}" for falha in falhas)


def mostrar_popup(mensagem, titulo="Instalacao"):
    """Mostra popup informativo nativo no Windows."""
    try:
        if platform.system() == "Windows":
            ctypes.windll.user32.MessageBoxW(None, str(mensagem), str(titulo), 0x40)
            return
    except Exception as exc:
        print(f"\nErro ao exibir popup: {exc}")
    print(f"\n=== {titulo} ===\n{mensagem}\n")


def main():
    print("=== CONFIGURADOR AUTOMATICO ===")
    print("Criando estrutura de pastas...")
    _inicializar_diretorios()
    _realocar_root_dir()

    sucesso, falhas = verificar_instalar_dependencias()

    if sucesso:
        mensagem = (
            "Todas as dependencias foram instaladas com sucesso!\n\n"
            "Agora voce pode executar o programa principal."
        )
        print(f"\n{mensagem}")
        mostrar_popup(mensagem, "Instalacao Completa")
        return

    lista_falhas = formatar_lista_falhas(falhas)
    mensagem = (
        "Houve problemas na instalacao das dependencias.\n\n"
        "Dependencias com falha:\n"
        f"{lista_falhas}\n\n"
        "Consulte as mensagens acima para corrigir."
    )

    print("\n" + "=" * 50)
    print("ERROS ENCONTRADOS:")
    print(lista_falhas)
    print("=" * 50)
    mostrar_popup(mensagem, "Erro na Instalacao")
    sys.exit(1)


if __name__ == "__main__":
    main()
