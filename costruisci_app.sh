#!/bin/zsh
# Costruisce l'applicazione con doppio clic (macOS: .app in dist/; su Windows usare costruisci_app.bat).
# Nome, versione e icona si cambiano in impostazioni_installer.txt.
# Uso:  ./costruisci_app.sh
set -e
cd "$(dirname "$0")"

IMPOSTAZIONI_OK=
eval "$(.venv/bin/python leggi_impostazioni.py sh)"
if [ -z "$IMPOSTAZIONI_OK" ]; then
  echo "Non riesco a leggere impostazioni_installer.txt (vedi il messaggio qui sopra)."; exit 1
fi
OPZ_ICONA=()
if [ -n "$ICONA" ]; then OPZ_ICONA=(--icon "$ICONA"); fi
echo "Programma: $NOME (versione $VERSIONE)"
if [ -n "$ICONA" ]; then echo "Icona: $ICONA"; else echo "Icona: predefinita"; fi

.venv/bin/pip install -q pyinstaller
.venv/bin/pyinstaller --noconfirm --windowed --name "$NOME" "${OPZ_ICONA[@]}" \
  --collect-all customtkinter --collect-all tksheet --collect-all ortools --collect-all tzdata \
  --hidden-import orario.motore --hidden-import orario.export --hidden-import orario.export_excel \
  --hidden-import orario.export_pdf --hidden-import orario.export_comune \
  app.py
echo
echo "Fatto. L'applicazione è in: dist/$NOME.app"
