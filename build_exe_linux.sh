#!/bin/bash
echo "Construindo executável com PyInstaller..."
python3 -m PyInstaller --onedir -w --noupx \
  --hidden-import=requests \
  --hidden-import=cryptography \
  --hidden-import=openpyxl \
  --hidden-import=pywin32 \
  --hidden-import=tkhtmlview \
  --hidden-import=markdown \
  --hidden-import=plyer \
  --add-data "config:config" \
  --add-data "docs:docs" \
  --add-data "cert_path:cert_path" \
  --add-data "packs/0:packs/0" \
  --icon=config/icone.ico \
  --version-file=docs/version_file.txt \
  download_nfse.py

echo "Copiando LICENSE para a pasta dist..."
cp -f docs/LICENSE dist/

echo "Copiando instalador.py e requirements.txt para dist/download_nfse..."
cp -f 1.instalador.py dist/download_nfse/
cp -f requirements.txt dist/download_nfse/

echo "Concluído! Executável criado em: dist/download_nfse/"
read -p "Pressione Enter para continuar..."