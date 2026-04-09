@echo off
setlocal
cd /d "%~dp0"

echo Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
  echo [ERRO] Python nao encontrado no PATH.
  exit /b 1
)

echo Instalando dependencias do projeto...
python -m pip install --upgrade pip
if errorlevel 1 (
  echo [ERRO] Falha ao atualizar o pip.
  exit /b 1
)

python -m pip install -r requirements.txt
if errorlevel 1 (
  echo [ERRO] Falha ao instalar as dependencias do projeto.
  exit /b 1
)

echo Instalando PyInstaller...
python -m pip install pyinstaller
if errorlevel 1 (
  echo [ERRO] Falha ao instalar o PyInstaller.
  exit /b 1
)

echo Limpando artefatos anteriores...
if exist build rmdir /s /q build
if exist dist\download_nfse rmdir /s /q dist\download_nfse
if exist download_nfse.spec del /q download_nfse.spec

echo Construindo executavel Qt com PyInstaller...
python -m PyInstaller --noconfirm --clean --onedir -w --noupx ^
  --name download_nfse ^
  --hidden-import=requests ^
  --hidden-import=cryptography ^
  --hidden-import=openpyxl ^
  --hidden-import=pywin32 ^
  --hidden-import=PySide6 ^
  --collect-all config ^
  --collect-all docs ^
  --collect-all cert_path ^
  --collect-all packs ^
  --icon=config/icone.ico ^
  --version-file=docs/version_file.txt ^
  download_nfse_qt.py

if errorlevel 1 (
  echo [ERRO] Falha durante o build do executavel.
  exit /b 1
)

echo Copiando arquivos auxiliares...
if exist docs\LICENSE copy /Y docs\LICENSE dist\ >nul
copy /Y README.md dist\download_nfse\README.md >nul

echo Build concluido com sucesso.
echo Executavel disponivel em: dist\download_nfse\
endlocal
