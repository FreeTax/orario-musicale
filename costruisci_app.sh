#!/bin/zsh
# Costruisce l'applicazione con doppio clic (macOS: .app in dist/; su Windows usare costruisci_app.bat).
# Uso:  ./costruisci_app.sh
set -e
cd "$(dirname "$0")"
.venv/bin/pip install -q pyinstaller
.venv/bin/pyinstaller --noconfirm --windowed --name "Orario Musicale" \
  --collect-all customtkinter --collect-all tksheet --collect-all ortools --collect-all tzdata \
  --hidden-import orario.motore --hidden-import orario.export --hidden-import orario.export_excel \
  --hidden-import orario.export_pdf --hidden-import orario.export_comune \
  app.py
echo
echo "Fatto. L'applicazione è in: dist/Orario Musicale.app"
