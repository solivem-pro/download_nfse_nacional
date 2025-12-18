@echo off
echo Construindo executável com PyInstaller...

python -m PyInstaller --onedir -w --noupx ^
  --hidden-import=requests ^
  --hidden-import=cryptography ^
  --hidden-import=openpyxl ^
  --hidden-import=pywin32 ^
  --hidden-import=tkhtmlview ^
  --hidden-import=markdown ^
  --hidden-import=plyer ^
  --collect-all config ^
  --collect-all docs ^
  --collect-all cert_path ^
  --collect-all packs ^
  --icon=config/icone.ico ^
  --version-file=docs/version_file.txt ^
  download_nfse.py

echo Copiando LICENSE para a pasta dist...
copy /Y docs\LICENSE dist\

echo Copiando instalador.py e requirements.txt para dist\download_nfse...
copy /Y 1.instalador.py dist\download_nfse\
copy /Y README.md dist\download_nfse\

echo Concluido! Executavel criado em: dist\download_nfse\
pause