"""Costruzione del file di input Excel (vuoto o precompilato): fogli, intestazioni, menu a tendina, istruzioni."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .costanti import FASCE, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_LMI, FOGLIO_PARAMETRI, FOGLIO_STUDENTI

HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
EX_FILL = PatternFill("solid", fgColor="F2F2F2")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True)
ESEMPIO = "ESEMPIO – cancellare"

COLONNE_STUDENTI = ["Classe", "Cognome", "Nome",
                    "Strumento 1", "Docente 1", "Strumento 2", "Docente 2",
                    "Ore consecutive (SI/NO)", "Giorno unico (SI/NO)", "Giorni NON disponibili",
                    "Comune", "Indirizzo", "Civico", "KM", "Note"]
LARGHEZZE_STUDENTI = {"Classe": 7, "Cognome": 22, "Nome": 22, "Strumento 1": 15, "Docente 1": 16,
                      "Strumento 2": 15, "Docente 2": 16, "Ore consecutive (SI/NO)": 13,
                      "Giorno unico (SI/NO)": 13, "Giorni NON disponibili": 17, "Comune": 20,
                      "Indirizzo": 26, "Civico": 8, "KM": 7, "Note": 30}
COLONNE_DOCENTI = ["Docente", "Strumento/i", "Aula", "Ore accompagnamento"] + FASCE + ["Note"]
COLONNE_GRUPPI = ["Gruppo", "Docente", "Studente 1", "Studente 2", "Studente 3", "Studente 4", "Studente 5", "Note"]
COLONNE_LMI = ["Laboratorio", "Classi", "Docente", "Aula", "Giorno e ora (mattino)",
               "Studenti (cognomi separati da virgola)", "Note"]
COLONNE_PARAMETRI = ["Parametro", "Valore", "Spiegazione"]
PARAMETRI_DEFAULT = [
    ("Max rientri per ragazzo", 2, "Numero massimo di pomeriggi a settimana (di norma). Chi abita vicino può arrivare al valore sotto."),
    ("Max rientri per chi abita vicino", 3, "Consentito solo ai ragazzi con KM sotto la soglia indicata alla riga seguente."),
    ("Soglia KM 'abita vicino'", 5, "Sotto questa distanza si accettano fino a 3 rientri."),
    ("Tempo massimo di calcolo (secondi)", 120, "Oltre questo tempo il programma restituisce la migliore soluzione trovata."),
    ("Indirizzo della scuola", "", "Via, numero e comune della scuola: punto di partenza per il calcolo dei mezzi pubblici (menu Orario → Aggiorna trasporti)."),
    ("Soglia minuti 'abita vicino'", 25, "Se ci sono i dati dei trasporti: sotto questi minuti di ritorno a casa il ragazzo è considerato vicino."),
]

ISTRUZIONI = [
    "ORARIO POMERIDIANO LICEO MUSICALE – FILE DI INPUT",
    "",
    "Questo file è l'unico input del programma. Compila i fogli seguenti, poi nel programma premi 'Calcola orario'.",
    "Le righe grigie con 'ESEMPIO' servono solo a far vedere come si compila: vanno cancellate.",
    "",
    "FOGLIO 'Studenti' – un rigo per ragazzo.",
    "  Classe: 1–5.  Docente 1 / Docente 2: devono coincidere con un nome del foglio 'Docenti' (menu a tendina).",
    "  Per la classe 5 il 2° strumento resta vuoto.",
    "  Ore consecutive = SI se le 2 ore di 1° strumento devono essere una di seguito all'altra (solo classi 1, 2, 5).",
    "  Giorno unico = SI se il ragazzo deve rientrare un solo giorno a settimana.",
    "  Giorni NON disponibili: es. 'Mar, Gio' se il ragazzo non può venire quei pomeriggi. Vuoto = tutti i giorni possibili.",
    "  Comune, Indirizzo, Civico: per calcolare i tempi di ritorno a casa con i mezzi (menu Orario → Aggiorna trasporti).",
    "  KM: distanza casa-scuola, usata quando mancano i dati dei mezzi. Senza né KM né mezzi si è trattati come vicini.",
    "",
    "FOGLIO 'Docenti' – un rigo per docente, con la disponibilità oraria.",
    "  Aula: aula fissa del docente, scritta in testa alla sua colonna nell'orario.",
    "  Colonne Lun 13:30 … Ven 16:30: scrivi X nelle ore in cui il docente È disponibile. Vuoto = non disponibile.",
    "  Ore accompagnamento: quante ore a settimana il docente (di pianoforte) fa da pianista accompagnatore. 0 se nessuna.",
    "  Il programma le colloca da solo in coda alle lezioni del docente e nell'orario scrive 'Pianista accomp.'.",
    "  Per fissare tu una di queste ore in una fascia precisa, scrivi A al posto della X (solo nelle ultime fasce del giorno).",
    "  Tutte le lezioni del docente (strumento e musica da camera) usano la stessa disponibilità.",
    "",
    "FOGLIO 'Gruppi LMC' – un rigo per gruppo di musica da camera (classi 3, 4, 5), da 2 a 5 ragazzi.",
    "  Scrivi i cognomi come nel foglio 'Studenti' (se due ragazzi hanno lo stesso cognome, aggiungi il nome).",
    "",
    "FOGLIO 'LMI' – laboratori di musica d'insieme (2 ore, orario del MATTINO). Un rigo per laboratorio.",
    "  Il programma NON li calcola: li ricopia così come sono in una pagina dell'orario.",
    "",
    "FOGLIO 'Parametri' – poche regole generali (max rientri, tempo di calcolo). Di norma non serve toccarlo.",
    "",
    "REGOLE FISSE (non modificabili da qui):",
    "  Classi 1 e 2: 2 ore 1° strumento + 1 ora 2° strumento.",
    "  Classi 3 e 4: 1 ora 1° strumento + 1 ora 2° strumento + 1 ora musica da camera.",
    "  Classe 5:     2 ore 1° strumento + 1 ora musica da camera.",
    "  Fasce: 13:30-14:30, 14:30-15:30, 15:30-16:30, 16:30-17:30, da lunedì a venerdì.",
]


def _intesta(ws, colonne: list[str], larghezze: dict[int, float] | None = None) -> None:
    ws.append(colonne)
    for i, _ in enumerate(colonne, start=1):
        c = ws.cell(row=1, column=i)
        c.font = BOLD
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = (larghezze or {}).get(i, 14)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32


def _grigia(ws, riga: int, n_col: int) -> None:
    for c in range(1, n_col + 1):
        ws.cell(row=riga, column=c).fill = EX_FILL


def costruisci_workbook(studenti: Iterable[dict] = (), docenti: Iterable[dict] = (), esempi: bool = True) -> Workbook:
    """Crea il workbook del file di input.

    studenti: dict con classe, cognome, nome, km, strum1, doc1, strum2, doc2 (km/doc possono essere vuoti).
    docenti:  dict con nome, strumenti, aula, note.
    esempi:   aggiunge righe grigie 'ESEMPIO' nei fogli vuoti.
    """
    studenti = list(studenti)
    docenti = list(docenti)
    wb = Workbook()

    ws = wb.active
    ws.title = "Istruzioni"
    ws.column_dimensions["A"].width = 115
    for r in ISTRUZIONI:
        ws.append([r])
    ws["A1"].font = Font(bold=True, size=14)
    for r, testo in enumerate(ISTRUZIONI, start=1):
        if testo.startswith("FOGLIO") or testo.startswith("REGOLE"):
            ws.cell(row=r, column=1).font = BOLD

    # Studenti
    ws = wb.create_sheet(FOGLIO_STUDENTI)
    col = {nome: i for i, nome in enumerate(COLONNE_STUDENTI, start=1)}
    _intesta(ws, COLONNE_STUDENTI, {col[k]: v for k, v in LARGHEZZE_STUDENTI.items()})

    def riga_studente(s: dict) -> list:
        valori = {
            "Classe": s["classe"], "Cognome": s["cognome"], "Nome": s["nome"],
            "Strumento 1": s.get("strum1", ""), "Docente 1": s.get("doc1", ""),
            "Strumento 2": s.get("strum2", ""), "Docente 2": s.get("doc2", ""),
            "Ore consecutive (SI/NO)": s.get("ore_consecutive", "NO"),
            "Giorno unico (SI/NO)": s.get("giorno_unico", "NO"),
            "Giorni NON disponibili": s.get("giorni_non_disp", ""),
            "Comune": s.get("comune", ""), "Indirizzo": s.get("indirizzo", ""),
            "Civico": s.get("civico", ""), "KM": s.get("km"), "Note": s.get("note", ""),
        }
        return [valori[c] for c in COLONNE_STUDENTI]

    for s in sorted(studenti, key=lambda x: (x["classe"], x["cognome"], x["nome"])):
        ws.append(riga_studente(s))
        r = ws.max_row
        if s.get("km") is None and not s.get("indirizzo"):
            ws.cell(row=r, column=col["KM"]).fill = WARN_FILL
        if not s.get("doc1"):
            ws.cell(row=r, column=col["Docente 1"]).fill = WARN_FILL
        if s["classe"] != 5 and not s.get("doc2"):
            ws.cell(row=r, column=col["Docente 2"]).fill = WARN_FILL
    if not studenti and esempi:
        ws.append(riga_studente({
            "classe": 1, "cognome": "ROSSI", "nome": "MARIO", "strum1": "PIANOFORTE", "doc1": "Bianchi",
            "strum2": "VIOLINO", "doc2": "Verdi", "comune": "Sambuca Pistoiese",
            "indirizzo": "Località Bellavalle", "civico": "68", "km": 27.5, "note": ESEMPIO}))
        ws.append(riga_studente({
            "classe": 4, "cognome": "NERI", "nome": "ANNA", "strum1": "CANTO", "doc1": "Verdi",
            "strum2": "PIANOFORTE", "doc2": "Bianchi", "giorno_unico": "SI", "giorni_non_disp": "Mar",
            "comune": "Pistoia", "indirizzo": "Via Nazario Sauro", "civico": "283", "km": 0.5,
            "note": ESEMPIO}))
        _grigia(ws, 2, len(COLONNE_STUDENTI))
        _grigia(ws, 3, len(COLONNE_STUDENTI))
    n_stud = max(ws.max_row, 2)
    fondo = n_stud + 200

    def lettera(nome: str) -> str:
        return get_column_letter(col[nome])

    dv = DataValidation(type="list", formula1='"SI,NO"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{lettera('Ore consecutive (SI/NO)')}2:{lettera('Giorno unico (SI/NO)')}{fondo}")
    dv = DataValidation(type="list", formula1="=Docenti!$A$2:$A$80", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{lettera('Docente 1')}2:{lettera('Docente 1')}{fondo}")
    dv.add(f"{lettera('Docente 2')}2:{lettera('Docente 2')}{fondo}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5")
    ws.add_data_validation(dv)
    dv.add(f"{lettera('Classe')}2:{lettera('Classe')}{fondo}")

    # Docenti
    ws = wb.create_sheet(FOGLIO_DOCENTI)
    larg = {1: 18, 2: 22, 3: 10, 4: 12, len(COLONNE_DOCENTI): 30}
    for i in range(5, 5 + len(FASCE)):
        larg[i] = 6.5
    _intesta(ws, COLONNE_DOCENTI, larg)
    for i in range(5, 5 + len(FASCE)):
        ws.cell(row=1, column=i).alignment = Alignment(wrap_text=True, text_rotation=90,
                                                      horizontal="center", vertical="bottom")
    ws.row_dimensions[1].height = 60
    for d in docenti:
        ws.append([d["nome"], d.get("strumenti", ""), d.get("aula", ""), d.get("ore_accomp", 0)]
                  + [""] * len(FASCE) + [d.get("note", "")])
        if d.get("note"):
            ws.cell(row=ws.max_row, column=1).fill = WARN_FILL
    if not docenti and esempi:
        ws.append(["Bianchi", "PIANOFORTE", "M1", 2] + ["X"] * 8 + [""] * 12 + [ESEMPIO])
        ws.append(["Verdi", "VIOLINO, CANTO", "M2", 0] + [""] * 4 + ["X"] * 12 + [""] * 4 + [ESEMPIO])
        _grigia(ws, 2, len(COLONNE_DOCENTI))
        _grigia(ws, 3, len(COLONNE_DOCENTI))
    dv = DataValidation(type="list", formula1='"X,A"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E2:{get_column_letter(4 + len(FASCE))}120")
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="20")
    ws.add_data_validation(dv)
    dv.add("D2:D120")

    # Gruppi LMC
    ws = wb.create_sheet(FOGLIO_GRUPPI)
    _intesta(ws, COLONNE_GRUPPI, {1: 8, 2: 18, 3: 18, 4: 18, 5: 18, 6: 18, 7: 18, 8: 30})
    if esempi:
        ws.append([1, "Verdi", "Neri", "Rossi", "", "", "", ESEMPIO])
        _grigia(ws, 2, len(COLONNE_GRUPPI))
    dv = DataValidation(type="list", formula1="=Docenti!$A$2:$A$80", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("B2:B120")
    col_cognome = get_column_letter(COLONNE_STUDENTI.index("Cognome") + 1)
    dv = DataValidation(type="list", formula1=f"=Studenti!${col_cognome}$2:${col_cognome}${n_stud + 200}",
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C2:G120")

    # LMI
    ws = wb.create_sheet(FOGLIO_LMI)
    _intesta(ws, COLONNE_LMI, {1: 26, 2: 14, 3: 18, 4: 22, 5: 22, 6: 60, 7: 24})
    if esempi:
        ws.append(["LMI ARCHI", "2ª, 3ª", "Verdi", "Aula M2, 1° piano", "Martedì 4ª-5ª ora", "Rossi, Neri", ESEMPIO])
        _grigia(ws, 2, len(COLONNE_LMI))
    dv = DataValidation(type="list", formula1="=Docenti!$A$2:$A$80", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C2:C60")

    # Parametri
    ws = wb.create_sheet(FOGLIO_PARAMETRI)
    _intesta(ws, COLONNE_PARAMETRI, {1: 34, 2: 10, 3: 80})
    for riga in PARAMETRI_DEFAULT:
        ws.append(list(riga))

    return wb


def crea_nuovo_file(percorso: Path) -> Path:
    """Crea un file di input vuoto (con righe di esempio) e ne ritorna il percorso."""
    wb = costruisci_workbook(esempi=True)
    wb.save(percorso)
    return percorso
