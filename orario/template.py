"""Costruzione del file di input Excel (vuoto o precompilato): fogli, intestazioni, menu a tendina, istruzioni."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .costanti import (
    COLONNE_AULE, FASCE, FOGLIO_ABBINAMENTI, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_IMPEGNI, FOGLIO_LMI,
    FOGLIO_PARAMETRI, FOGLIO_STUDENTI, GIORNI, NOMI_TIPO, ORE,
)

HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
EX_FILL = PatternFill("solid", fgColor="F2F2F2")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True)
ESEMPIO = "ESEMPIO – cancellare"

COLONNE_STUDENTI = ["Classe", "Cognome", "Nome",
                    "Strumento 1", "Docente 1", "Strumento 2", "Docente 2",
                    "1° strumento attaccato (SI/NO)", "Giorno unico (SI/NO)", "Giorni NON disponibili",
                    "Comune", "Indirizzo", "Civico", "KM", "Minuti per tornare a casa", "Note"]
LARGHEZZE_STUDENTI = {"Classe": 7, "Cognome": 22, "Nome": 22, "Strumento 1": 15, "Docente 1": 16,
                      "Strumento 2": 15, "Docente 2": 16, "1° strumento attaccato (SI/NO)": 17,
                      "Giorno unico (SI/NO)": 13, "Giorni NON disponibili": 17, "Comune": 20,
                      "Indirizzo": 26, "Civico": 8, "KM": 7, "Minuti per tornare a casa": 16, "Note": 30}
COLONNE_DOCENTI = (["Docente", "Strumento/i"] + COLONNE_AULE + ["Ore accompagnamento"]
                   + FASCE + ["Ore dichiarate", "Note"])
RIGA_TOTALE = "TOTALE"   # ultima riga del foglio Docenti: somme, non è un docente
COLONNE_GRUPPI = ["Gruppo", "Docente", "Studente 1", "Studente 2", "Studente 3", "Studente 4", "Studente 5", "Note"]
COLONNE_LMI = ["Laboratorio", "Classi", "Docente", "Aula", "Giorno e ora (mattino)",
               "Studenti (Cognome Nome, separati da virgola)", "Note"]
COLONNE_IMPEGNI = ["Classe", "Cognome", "Nome"] + FASCE + ["Note"]
COLONNE_ABBINAMENTI = ["Docente", "Studente", "Tipo di lezione", "Giorno", "Ora", "Note"]
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
    "  1° strumento attaccato (solo classi 1, 2, 5, che hanno 2 ore): SI = le due ore una di seguito all'altra.",
    "  NO o casella vuota = in due giorni con almeno un giorno in mezzo (lun/mer, mar/ven...), mai di fila.",
    "  Giorno unico = SI se il ragazzo deve rientrare un solo giorno a settimana.",
    "  Giorni NON disponibili: es. 'Mar, Gio' se il ragazzo non può venire quei pomeriggi. Vuoto = tutti i giorni possibili.",
    "  Comune, Indirizzo, Civico: per calcolare i tempi di ritorno a casa con i mezzi (menu Orario → Aggiorna trasporti).",
    "  Minuti per tornare a casa: lo scrive il programma nelle caselle vuote (il migliore delle 4 fasce);",
    "  il dettaglio ora per ora è nel foglio 'Trasporti'. Si può correggere a mano: il numero scritto vince",
    "  su quello calcolato e non viene più sovrascritto. Per farlo ricalcolare, svuotare la casella.",
    "  KM: distanza casa-scuola, usata quando mancano i dati dei mezzi. Senza né KM né mezzi si è trattati come vicini.",
    "",
    "FOGLIO 'Docenti' – un rigo per docente, con la disponibilità oraria.",
    "  Nel programma il pulsante 'Compila dagli studenti' aggiunge da solo i docenti che compaiono nel foglio",
    "  Studenti e nei gruppi, con i loro strumenti: restano da mettere a mano solo aule e disponibilità.",
    "  Aula Lun … Aula Ven: l'aula del docente in ciascun giorno (se è sempre la stessa, si ripete).",
    "  Viene scritta in testa alla sua colonna nell'orario del giorno.",
    "  Colonne Lun 13:30 … Ven 16:30: scrivi X nelle ore in cui il docente È disponibile. Vuoto = non disponibile.",
    "  Ore accompagnamento: quante ore a settimana il docente (di pianoforte) fa da pianista accompagnatore. 0 se nessuna.",
    "  Il programma le colloca da solo in coda alle lezioni del docente e nell'orario scrive 'Pianista accomp.'.",
    "  Per fissare tu una di queste ore in una fascia precisa, scrivi A al posto della X (solo nelle ultime fasce del giorno).",
    "  Tutte le lezioni del docente (strumento e musica da camera) usano la stessa disponibilità.",
    "",
    "FOGLIO 'Gruppi LMC' – un rigo per gruppo di musica da camera (classi 3, 4, 5), da 2 a 5 ragazzi.",
    "  Scrivi 'Cognome Nome' come nel foglio 'Studenti'. Nel programma c'è il menu a tendina con tutti i ragazzi:",
    "  se scrivi solo il cognome e non ci sono omonimi, il nome viene aggiunto da solo.",
    "",
    "FOGLIO 'Impegni studenti' – le ore in cui un ragazzo NON può venire (sport, catechismo, altro).",
    "  ATTENZIONE: qui è al contrario del foglio Docenti. La X segna l'ora in cui il ragazzo NON c'è.",
    "  Chi non ha impegni si lascia con la riga vuota. Nel programma il pulsante 'Compila dagli studenti'",
    "  ricopia l'elenco dei nomi dal foglio Studenti.",
    "",
    "FOGLIO 'Abbinamenti fissi' – lezioni già decise, che il programma deve rispettare così come sono.",
    "  Una riga per lezione: docente, studente, tipo (1° strumento, 2° strumento, musica da camera), giorno e ora.",
    "  Il tipo si può lasciare vuoto: viene dedotto dal docente. Serve per bloccare gli incastri già concordati.",
    "  Nella colonna Studente, al posto di un ragazzo, si può scrivere 'Gruppo 5' per fissare l'ora di un gruppo",
    "  di musica da camera, oppure il nome di un LABORATORIO del foglio LMI (tipo 'Laboratorio LMI'):",
    "  in quell'ora vengono occupati il docente e tutti i ragazzi del gruppo o del laboratorio.",
    "  Il docente si può lasciare vuoto: si prende da 'Gruppi LMC' o da 'LMI'.",
    "  Le 2 ore di 1° strumento non sono per forza attaccate: se le vuoi di seguito, mettile qui in due righe",
    "  su due ore consecutive (oppure segna '1° strumento attaccato = SI' nel foglio Studenti).",
    "",
    "FOGLIO 'LMI' – laboratori di musica d'insieme (2 ore, orario del MATTINO). Un rigo per laboratorio.",
    "  Il programma NON li calcola: li ricopia così come sono in una pagina dell'orario. Se però un",
    "  laboratorio si tiene nel pomeriggio, lo si mette nel foglio 'Abbinamenti fissi' con giorno e ora.",
    "  Studenti: 'Cognome Nome' separati da virgola. Anche qui il nome viene aggiunto da solo quando non ci sono omonimi.",
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

    col_cognome = get_column_letter(COLONNE_STUDENTI.index("Cognome") + 1)

    # Studenti
    ws = wb.create_sheet(FOGLIO_STUDENTI)
    col = {nome: i for i, nome in enumerate(COLONNE_STUDENTI, start=1)}
    _intesta(ws, COLONNE_STUDENTI, {col[k]: v for k, v in LARGHEZZE_STUDENTI.items()})

    def riga_studente(s: dict) -> list:
        valori = {
            "Classe": s["classe"], "Cognome": s["cognome"], "Nome": s["nome"],
            "Strumento 1": s.get("strum1", ""), "Docente 1": s.get("doc1", ""),
            "Strumento 2": s.get("strum2", ""), "Docente 2": s.get("doc2", ""),
            "1° strumento attaccato (SI/NO)": s.get("primo_attaccato", ""),
            "Giorno unico (SI/NO)": s.get("giorno_unico", "NO"),
            "Giorni NON disponibili": s.get("giorni_non_disp", ""),
            "Comune": s.get("comune", ""), "Indirizzo": s.get("indirizzo", ""),
            "Civico": s.get("civico", ""), "KM": s.get("km"),
            "Minuti per tornare a casa": "",   # lo scrive «Aggiorna trasporti»
            "Note": s.get("note", ""),
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
    dv.add(f"{lettera(COLONNE_STUDENTI[7])}2:{lettera('Giorno unico (SI/NO)')}{fondo}")
    dv = DataValidation(type="list", formula1="=Docenti!$A$2:$A$80", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{lettera('Docente 1')}2:{lettera('Docente 1')}{fondo}")
    dv.add(f"{lettera('Docente 2')}2:{lettera('Docente 2')}{fondo}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5")
    ws.add_data_validation(dv)
    dv.add(f"{lettera('Classe')}2:{lettera('Classe')}{fondo}")

    # Impegni degli studenti: al contrario dei docenti, la casella segnata è un'ora in cui NON ci sono
    ws = wb.create_sheet(FOGLIO_IMPEGNI)
    coli = {nome: i for i, nome in enumerate(COLONNE_IMPEGNI, start=1)}
    prima_i = coli[FASCE[0]]
    largi = {coli["Classe"]: 7, coli["Cognome"]: 22, coli["Nome"]: 22, len(COLONNE_IMPEGNI): 30}
    for i in range(prima_i, prima_i + len(FASCE)):
        largi[i] = 6.5
    _intesta(ws, COLONNE_IMPEGNI, largi)
    for i in range(prima_i, prima_i + len(FASCE)):
        ws.cell(row=1, column=i).alignment = Alignment(wrap_text=True, text_rotation=90,
                                                      horizontal="center", vertical="bottom")
    ws.row_dimensions[1].height = 60
    for s_ in sorted(studenti, key=lambda x: (x["classe"], x["cognome"], x["nome"])):
        ws.append([s_["classe"], s_["cognome"], s_["nome"]] + [""] * len(FASCE) + [""])
    if not studenti and esempi:
        ws.append([4, "NERI", "ANNA"] + ["X"] * 4 + [""] * 16 + [ESEMPIO + " (il martedì fa sport)"])
        _grigia(ws, 2, len(COLONNE_IMPEGNI))
    dv = DataValidation(type="list", formula1='"X"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(prima_i)}2:{get_column_letter(prima_i + len(FASCE) - 1)}{max(ws.max_row, 2) + 200}")

    # Docenti
    ws = wb.create_sheet(FOGLIO_DOCENTI)
    cold = {nome: i for i, nome in enumerate(COLONNE_DOCENTI, start=1)}
    prima_fascia = cold[FASCE[0]]
    larg = {cold["Docente"]: 18, cold["Strumento/i"]: 22, cold["Ore accompagnamento"]: 12,
            cold["Ore dichiarate"]: 10, len(COLONNE_DOCENTI): 30}
    for nome in COLONNE_AULE:
        larg[cold[nome]] = 9
    for i in range(prima_fascia, prima_fascia + len(FASCE)):
        larg[i] = 6.5
    _intesta(ws, COLONNE_DOCENTI, larg)
    for i in range(prima_fascia, prima_fascia + len(FASCE)):
        ws.cell(row=1, column=i).alignment = Alignment(wrap_text=True, text_rotation=90,
                                                      horizontal="center", vertical="bottom")
    ws.row_dimensions[1].height = 60
    for d in docenti:
        aule = d.get("aule") or [d.get("aula", "")] * len(COLONNE_AULE)
        ws.append([d["nome"], d.get("strumenti", "")] + list(aule)[:len(COLONNE_AULE)]
                  + [d.get("ore_accomp", 0)] + [""] * len(FASCE) + ["", d.get("note", "")])
        if d.get("note"):
            ws.cell(row=ws.max_row, column=1).fill = WARN_FILL
    if not docenti and esempi:
        ws.append(["Bianchi", "PIANOFORTE"] + ["M1"] * 5 + [2] + ["X"] * 8 + [""] * 12 + ["", ESEMPIO])
        ws.append(["Verdi", "VIOLINO, CANTO", "M2", "M2", "M3", "M3", "M2", 0]
                  + [""] * 4 + ["X"] * 12 + [""] * 4 + ["", ESEMPIO])
        _grigia(ws, 2, len(COLONNE_DOCENTI))
        _grigia(ws, 3, len(COLONNE_DOCENTI))
    dv = DataValidation(type="list", formula1='"X,A"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(prima_fascia)}2:{get_column_letter(prima_fascia + len(FASCE) - 1)}120")
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="20")
    ws.add_data_validation(dv)
    lettera_acc = get_column_letter(cold["Ore accompagnamento"])
    dv.add(f"{lettera_acc}2:{lettera_acc}120")

    # Gruppi LMC
    ws = wb.create_sheet(FOGLIO_GRUPPI)
    _intesta(ws, COLONNE_GRUPPI, {1: 8, 2: 18, 3: 18, 4: 18, 5: 18, 6: 18, 7: 18, 8: 30})
    if esempi:
        ws.append([1, "Verdi", "Neri", "Rossi", "", "", "", ESEMPIO])
        _grigia(ws, 2, len(COLONNE_GRUPPI))
    dv = DataValidation(type="list", formula1="=Docenti!$A$2:$A$80", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("B2:B120")
    dv = DataValidation(type="list", formula1=f"=Studenti!${col_cognome}$2:${col_cognome}${n_stud + 200}",
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("C2:G120")

    # Abbinamenti fissi: lezioni già decise, che il programma deve rispettare
    ws = wb.create_sheet(FOGLIO_ABBINAMENTI)
    _intesta(ws, COLONNE_ABBINAMENTI, {1: 22, 2: 34, 3: 20, 4: 14, 5: 10, 6: 44})
    if esempi:
        ws.append(["Verdi", "Neri Anna", "2° strumento", "Mercoledì", "14:30", ESEMPIO])
        _grigia(ws, 2, len(COLONNE_ABBINAMENTI))
    fondo_ab = max(ws.max_row, 2) + 200
    for colonna, valori in (("A", "=Docenti!$A$2:$A$80"),
                            ("B", f"=Studenti!${col_cognome}$2:${col_cognome}${n_stud + 200}")):
        # solo un suggerimento: nella colonna Studente si può scrivere anche un laboratorio LMI
        dv = DataValidation(type="list", formula1=valori, allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dv)
        dv.add(f"{colonna}2:{colonna}{fondo_ab}")
    for colonna, valori in (("C", list(NOMI_TIPO)), ("D", list(GIORNI)), ("E", list(ORE))):
        dv = DataValidation(type="list", formula1='"' + ",".join(valori) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{colonna}2:{colonna}{fondo_ab}")

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


# ── Riallineamento del formato di un file già esistente ──────────────────────

def rinfresca_formato(percorso: Path) -> list[str]:
    """Riporta un file già compilato al formato di oggi, senza toccare i dati.

    Riscrive il foglio Istruzioni e rimette, foglio per foglio, le larghezze delle colonne,
    lo stile dell'intestazione e i menu a tendina, che nei file più vecchi mancano sulle
    colonne aggiunte dopo. I fogli con colonne diverse dal modello vengono lasciati stare.
    """
    from copy import copy

    from openpyxl import load_workbook
    from openpyxl.utils import range_boundaries

    from .costanti import FOGLI_DATI
    from .lettura import _salva_atomico

    wb = load_workbook(percorso)
    # il modello si costruisce con lo stesso numero di righe, così i menu coprono tutte le righe vere
    n_stud = max(0, (wb[FOGLIO_STUDENTI].max_row - 1) if FOGLIO_STUDENTI in wb.sheetnames else 0)
    n_doc = max(0, (wb[FOGLIO_DOCENTI].max_row - 1) if FOGLIO_DOCENTI in wb.sheetnames else 0)
    modello = costruisci_workbook(
        studenti=[{"classe": 1, "cognome": "", "nome": ""} for _ in range(n_stud)],
        docenti=[{"nome": ""} for _ in range(n_doc)], esempi=False)

    fatti: list[str] = []
    if "Istruzioni" in wb.sheetnames:
        del wb["Istruzioni"]
    ws = wb.create_sheet("Istruzioni", 0)
    ws.column_dimensions["A"].width = 115
    for testo in ISTRUZIONI:
        ws.append([testo])
    ws["A1"].font = Font(bold=True, size=14)
    for r, testo in enumerate(ISTRUZIONI, start=1):
        if testo.startswith(("FOGLIO", "REGOLE")):
            ws.cell(row=r, column=1).font = BOLD
    fatti.append("foglio «Istruzioni» riscritto")

    for nome in FOGLI_DATI:
        if nome not in wb.sheetnames or nome not in modello.sheetnames:
            continue
        src, dst = modello[nome], wb[nome]
        intest_src = [c.value for c in src[1]]
        if [c.value for c in dst[1]] != intest_src:
            continue                      # colonne diverse dal modello: meglio non toccare niente
        for i in range(1, len(intest_src) + 1):
            lettera_col = get_column_letter(i)
            larghezza = src.column_dimensions[lettera_col].width
            if larghezza:
                dst.column_dimensions[lettera_col].width = larghezza
            testa_src, testa_dst = src.cell(1, i), dst.cell(1, i)
            testa_dst.font = copy(testa_src.font)
            testa_dst.fill = copy(testa_src.fill)
            testa_dst.alignment = copy(testa_src.alignment)
        if src.row_dimensions[1].height:
            dst.row_dimensions[1].height = src.row_dimensions[1].height
        dst.freeze_panes = src.freeze_panes
        # menu a tendina: si rifanno tutti, allungati fino in fondo alle righe vere
        dst.data_validations.dataValidation = []
        fondo = max(dst.max_row, 2) + 200
        for dv in src.data_validations.dataValidation:
            nuovo = DataValidation(type=dv.type, operator=dv.operator, formula1=dv.formula1,
                                   formula2=dv.formula2, allow_blank=dv.allowBlank,
                                   showErrorMessage=dv.showErrorMessage)
            dst.add_data_validation(nuovo)
            for rif in dv.sqref.ranges:
                c1, _, c2, _ = range_boundaries(str(rif))
                nuovo.add(f"{get_column_letter(c1)}2:{get_column_letter(c2)}{fondo}")
        fatti.append(f"larghezze e menu a tendina del foglio «{nome}»")
    _salva_atomico(wb, percorso)
    return fatti
