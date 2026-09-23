"""Legge impostazioni_installer.txt per gli script di costruzione.

Uso (non serve lanciarlo a mano, lo fanno i .bat e lo .sh):
    python leggi_impostazioni.py bat   → righe "set ..." per i file .bat
    python leggi_impostazioni.py sh    → righe VAR='...' per la shell del Mac

In entrambi i casi scrive anche installer/impostazioni_generate.iss, che installer.iss include.
Se il file delle impostazioni manca o una voce è vuota, usa i valori predefiniti.
Se l'icona indicata non esiste, avvisa e va avanti senza icona.
"""

import os
import re
import sys

CARTELLA = os.path.dirname(os.path.abspath(__file__))
FILE_IMPOSTAZIONI = os.path.join(CARTELLA, "impostazioni_installer.txt")
FILE_ISS = os.path.join(CARTELLA, "installer", "impostazioni_generate.iss")

PREDEFINITE = {"nome": "Orario Musicale", "versione": "1.0", "icona": "", "nome_installer": ""}


def avviso(testo):
    print("  ATTENZIONE: " + testo, file=sys.stderr)


def leggi():
    valori = dict(PREDEFINITE)
    if not os.path.exists(FILE_IMPOSTAZIONI):
        avviso("impostazioni_installer.txt non trovato: uso nome 'Orario Musicale', versione 1.0, senza icona.")
        return valori
    with open(FILE_IMPOSTAZIONI, encoding="utf-8-sig") as f:   # utf-8-sig toglie il BOM di Blocco note
        for numero, riga in enumerate(f, 1):
            riga = riga.strip()
            if not riga or riga.startswith("#"):
                continue
            if "=" not in riga:
                avviso(f"riga {numero} di impostazioni_installer.txt ignorata (manca '='): {riga}")
                continue
            chiave, valore = riga.split("=", 1)
            chiave = chiave.strip().lower()
            valore = valore.strip().strip('"').strip("'")
            if chiave not in PREDEFINITE:
                avviso(f"riga {numero}: voce sconosciuta '{chiave}', ignorata.")
                continue
            if valore:
                valori[chiave] = valore
    return valori


def formato_immagine(percorso):
    """Riconosce il formato dai primi byte: 'ico', 'png', 'icns', 'jpeg', 'bmp' o None."""
    with open(percorso, "rb") as f:
        testa = f.read(8)
    if testa[:4] == b"\x00\x00\x01\x00":
        return "ico"
    if testa[:4] == b"\x89PNG":
        return "png"
    if testa[:4] == b"icns":
        return "icns"
    if testa[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if testa[:2] == b"BM":
        return "bmp"
    return None


def dimensioni_ico(percorso):
    """Lati delle immagini contenute in un .ico (0 nell'intestazione significa 256)."""
    import struct
    with open(percorso, "rb") as f:
        dati = f.read(6 + 16 * 64)
    quante = struct.unpack_from("<H", dati, 4)[0]
    lati = set()
    for i in range(min(quante, 64)):
        if 6 + 16 * i + 2 > len(dati):
            break
        larghezza, altezza = dati[6 + 16 * i], dati[6 + 16 * i + 1]
        lati.add(max(larghezza or 256, altezza or 256))
    return lati


def sistema(valori):
    nome = valori["nome"]
    if re.search(r'[\\/:*?"<>|]', nome):
        avviso("il nome contiene caratteri non ammessi nei nomi di file (\\ / : * ? \" < > |): li tolgo.")
        nome = re.sub(r'[\\/:*?"<>|]', "", nome).strip() or PREDEFINITE["nome"]
    valori["nome"] = nome

    icona = valori["icona"]
    if icona:
        percorso = icona if os.path.isabs(icona) else os.path.join(CARTELLA, icona)
        percorso = os.path.normpath(percorso)
        if not os.path.exists(percorso):
            avviso(f"icona '{icona}' non trovata: il programma avrà l'icona predefinita.")
            icona = ""
        else:
            estensione = os.path.splitext(percorso)[1].lower()
            if estensione not in (".ico", ".icns", ".png"):
                avviso(f"l'icona è un file {estensione}: meglio un .ico (Windows) o .icns (Mac).")
            formato = formato_immagine(percorso)
            if estensione == ".ico" and formato != "ico":
                avviso(f"'{icona}' ha estensione .ico ma dentro è un file {formato or 'sconosciuto'}: "
                       "rinominare non basta. Va convertito in .ico vero (per esempio con un "
                       "convertitore online 'PNG to ICO'). Vado avanti senza icona.")
                percorso = ""
            elif formato == "ico":
                dimensioni = dimensioni_ico(percorso)
                if dimensioni and max(dimensioni) < 48:
                    avviso(f"l'icona contiene solo {sorted(dimensioni)} pixel: sul desktop verrà sgranata. "
                           "Meglio un .ico con anche 48 e 256 pixel.")
            icona = percorso
    valori["icona"] = icona

    nome_installer = valori["nome_installer"]
    if nome_installer.lower().endswith(".exe"):
        nome_installer = nome_installer[:-4]
    if not nome_installer:
        nome_installer = f"{nome.replace(' ', '')}-setup-{valori['versione']}"
    valori["nome_installer"] = re.sub(r'[\\/:*?"<>|]', "", nome_installer)
    return valori


def scrivi_iss(valori):
    def stringa(testo):
        return '"' + testo.replace('"', '""') + '"'

    righe = [
        "; File generato da leggi_impostazioni.py a partire da impostazioni_installer.txt.",
        "; Non modificarlo a mano: viene riscritto a ogni costruzione.",
        f"#define NomeApp {stringa(valori['nome'])}",
        f"#define Versione {stringa(valori['versione'])}",
        f"#define NomeInstaller {stringa(valori['nome_installer'])}",
        f"#define Icona {stringa(valori['icona'])}",
        "",
    ]
    os.makedirs(os.path.dirname(FILE_ISS), exist_ok=True)
    with open(FILE_ISS, "w", encoding="utf-8") as f:
        f.write("\n".join(righe))


def main():
    modo = sys.argv[1] if len(sys.argv) > 1 else "sh"
    if modo == "bat" and os.name == "nt":
        # cmd legge l'output con la codifica della console (es. cp850), non con quella di Windows.
        try:
            import ctypes
            cp = ctypes.windll.kernel32.GetConsoleOutputCP()
            sys.stdout.reconfigure(encoding=f"cp{cp}", errors="replace")
        except Exception:
            pass

    valori = sistema(leggi())
    scrivi_iss(valori)

    coppie = [
        ("NOME", valori["nome"]),
        ("VERSIONE", valori["versione"]),
        ("ICONA", valori["icona"]),
        ("NOME_INSTALLER", valori["nome_installer"]),
    ]
    if modo == "bat":
        for nome, valore in coppie:
            print(f'set "{nome}={valore}"')
        print('set "IMPOSTAZIONI_OK=1"')
    else:
        for nome, valore in coppie:
            print(f"{nome}='" + valore.replace("'", "'\\''") + "'")
        print("IMPOSTAZIONI_OK=1")


if __name__ == "__main__":
    main()
