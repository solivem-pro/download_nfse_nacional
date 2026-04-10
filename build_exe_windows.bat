@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "PYTHON_EXE=python"
set "PYTHON_ARGS="
python --version >nul 2>&1
if errorlevel 1 (
  py -3.12 --version >nul 2>&1
  if not errorlevel 1 (
    set "PYTHON_EXE=py"
    set "PYTHON_ARGS=-3.12"
  ) else (
    set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
    if not exist "!PYTHON_EXE!" (
      echo [ERRO] Python nao encontrado no PATH nem na instalacao padrao.
      exit /b 1
    )
  )
)

echo Instalando dependencias do projeto...
"%PYTHON_EXE%" %PYTHON_ARGS% -m pip install --upgrade pip
if errorlevel 1 (
  echo [ERRO] Falha ao atualizar o pip.
  exit /b 1
)

"%PYTHON_EXE%" %PYTHON_ARGS% -m pip install -r requirements.txt
if errorlevel 1 (
  echo [ERRO] Falha ao instalar as dependencias do projeto.
  exit /b 1
)

echo Instalando PyInstaller...
"%PYTHON_EXE%" %PYTHON_ARGS% -m pip install pyinstaller
if errorlevel 1 (
  echo [ERRO] Falha ao instalar o PyInstaller.
  exit /b 1
)

echo Limpando artefatos anteriores...
if exist build rmdir /s /q build
if exist dist\download_nfse.exe del /q dist\download_nfse.exe
if exist dist\download_nfse rmdir /s /q dist\download_nfse
if exist download_nfse.spec del /q download_nfse.spec

echo Construindo executavel unico com PyInstaller...
"%PYTHON_EXE%" %PYTHON_ARGS% -m PyInstaller --noconfirm --clean --onefile -w --noupx ^
  --name download_nfse ^
  --hidden-import=requests ^
  --hidden-import=cryptography ^
  --hidden-import=openpyxl ^
  --hidden-import=reportlab ^
  --hidden-import=PySide6 ^
  --add-data "config\cadastros.db;config" ^
  --add-data "config\cadastros.json;config" ^
  --add-data "config\config.json;config" ^
  --add-data "config\icone.ico;config" ^
  --add-data "docs;docs" ^
  --add-data "cert_path;cert_path" ^
  --add-data "packs;packs" ^
  --add-data "README.md;." ^
  --icon=config/icone.ico ^
  --version-file=docs/version_file.txt ^
  download_nfse_qt.py

if errorlevel 1 (
  echo [ERRO] Falha durante o build do executavel.
  exit /b 1
)

echo Build concluido com sucesso.
echo Executavel disponivel em: dist\download_nfse.exe
endlocal
