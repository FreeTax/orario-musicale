"""Importa il file Word "classi e gruppi" (e i recapiti) nel file di input del programma.

Legge dal documento Word:
  * le cinque tabelle delle classi  → studenti, strumenti e docenti
  * le tabelle riassuntive per strumento → elenco docenti, ore, ore di accompagnamento,
    cattedre non ancora assegnate
  * la tabella delle FORMAZIONI  → gruppi di musica da camera (LMC)
  * la tabella MUSICA D'INSIEME  → laboratori del mattino (LMI), che il programma ricopia
E dal file Excel dei recapiti: comune, indirizzo, civico e chilometri di ogni ragazzo.

Uso:
    .venv/bin/python importa_docx.py                        # cerca i file più recenti nella cartella
    .venv/bin/python importa_docx.py classi.docx recapiti.xls uscita.xlsx
"""

from __future__ import annotations

import difflib
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import docx
import xlrd
from docx.oxml.ns import qn

from orario.costanti import FOGLIO_LMI, FOGLIO_PARAMETRI
from orario.template import costruisci_workbook

ROOT = Path(__file__).parent
USCITA_DEFAULT = ROOT / "input_orario.xlsx"

# ── Nomi che nel Word compaiono in più forme, o cattedre non ancora assegnate ──
ALIAS_DOCENTI = {
    "PACINI M.": "Pacini M",
    "PACINI": "Pacini M",            # nelle tabelle chitarra e LMI è scritto senza iniziale
    "BONACCORSI": "Buonaccorsi",     # tromba: due grafie nello stesso documento
    "PERCUSSIONE 2": "Percussioni 2",
    "PERC 2": "Percussioni 2",
    "VIOLINO": "Violino 2",          # la seconda cattedra di violino, ancora senza nome
    "VIOLINO 2": "Violino 2",
}
# cattedre senza titolare: il nome va scritto a mano quando si saprà
DA_NOMINARE = {"Percussioni 2": "Cattedra Percussioni 2: docente da nominare",
               "Violino 2": "Cattedra Violino 2: docente da nominare"}

ALIAS_STRUMENTI = {
    "VIOLOCELLO": "VIOLONCELLO", "V.CELLO": "VIOLONCELLO", "CELLO": "VIOLONCELLO",
    "C.BASSO": "CONTRABBASSO", "CB": "CONTRABBASSO", "C. BASSO": "CONTRABBASSO",
    "PERCUSSIONE": "PERCUSSIONI", "PERC": "PERCUSSIONI",
    "CHIT": "CHITARRA", "PF": "PIANOFORTE", "VL": "VIOLINO", "VLA": "VIOLA",
    "FL": "FLAUTO", "OB": "OBOE", "TR": "TROMBA", "COR": "CORNO", "CL": "CLARINETTO",
    "SAX": "SAXOFONO", "SAXOFONO": "SAXOFONO",
}
# abbreviazioni usate accanto ai cognomi nei gruppi: non sono cognomi
SIGLE_STRUMENTO = {"PF", "VL", "VLA", "CELLO", "CB", "CHIT", "CANTO", "PERC", "SAX", "FL",
                   "OB", "TR", "COR", "CL", "VIOLINO", "VIOLA", "VIOLONCELLO", "CONTRABBASSO",
                   "CHITARRA", "PIANOFORTE", "FLAUTO", "OBOE", "TROMBA", "CORNO", "CLARINETTO",
                   "PERCUSSIONI", "SAXOFONO"}

# Aule dell'orario 2025-26: punto di partenza, da verificare
AULE_PRECEDENTI = {
    "Vitangeli": "1ALM", "Lala": "2° p A", "Bini": "5ALM", "Pacini M": "M0", "Donnini": "M1",
    "Bartolozzi": "M2", "Francini": "M3", "Salaris": "M4", "Paganin": "M5", "Corsini": "M6",
    "Pacini C": "4ALM", "Paccosi": "M7", "Simonelli": "M8", "Lissoni": "3ALM", "Innocenti": "3ALM",
    "Mazzei": "M9", "Percussioni 2": "M 10", "Di Bonaventura": "M 11", "Caselli": "2ALM",
    "Buonaccorsi": "2ALM",
}


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).replace("’", "'").upper()
    return re.sub(r"\s+", " ", s).strip()


def docente(nome: str) -> str:
    n = re.sub(r"\s+", " ", (nome or "").strip()).strip(".,;")
    return ALIAS_DOCENTI.get(norm(n), n)


def strumento(s: str) -> str:
    n = norm(s).strip(".,;")
    return ALIAS_STRUMENTI.get(n, n)


# ── Lettura delle tabelle Word ───────────────────────────────────────────────

def celle(row) -> list[str]:
    """Testo delle celle, senza le ripetizioni introdotte dalle celle unite."""
    out, visti = [], set()
    for c in row.cells:
        if id(c._tc) in visti:
            continue
        visti.add(id(c._tc))
        out.append(re.sub(r"\s*\n\s*", " / ", c.text).strip())
    return out


def celle_con_colonne(row) -> list[tuple[int, int, str]]:
    """(colonna iniziale, colonne occupate, testo) per ogni cella: serve dove le celle sono unite."""
    out, col = [], 0
    for tc in row._tr.findall(qn("w:tc")):
        span = tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
        n = int(span.get(qn("w:val"))) if span is not None else 1
        paragrafi = []
        for par in tc.findall(qn("w:p")):
            testo = "".join(x.text or "" for x in par.iter(qn("w:t"))).strip()
            if testo:
                paragrafi.append(testo)
        out.append((col, n, " / ".join(paragrafi)))
        col += n
    return out


@dataclass
class Studente:
    classe: int
    cognome: str
    nome: str
    strum1: str = ""
    doc1: str = ""
    strum2: str = ""
    doc2: str = ""
    km: float | None = None
    comune: str = ""
    indirizzo: str = ""
    civico: str = ""
    note: list[str] = field(default_factory=list)

    @property
    def chiave(self) -> tuple[str, str]:
        return norm(self.cognome), norm(self.nome)


def leggi_classi(doc) -> tuple[list[Studente], list[str]]:
    """Le cinque tabelle delle classi. Ritorna gli studenti e le segnalazioni."""
    studenti, note = [], []
    for classe in range(1, 6):
        t = doc.tables[classe - 1]
        for row in t.rows:
            c = celle(row)
            if len(c) < 4 or not c[1] or norm(c[1]) == "COGNOME" or norm(c[0]).startswith("CLASSE"):
                continue
            cognome = c[1].replace("*", "").strip()
            nome = c[2].replace("*", "").strip()
            if not cognome or not nome:
                continue
            s = Studente(classe=classe, cognome=cognome.upper(), nome=nome.upper(),
                         strum1=strumento(c[3]), doc1=docente(c[4]) if len(c) > 4 else "")
            if classe != 5:
                s.strum2 = strumento(c[5]) if len(c) > 5 else ""
                s.doc2 = docente(c[6]) if len(c) > 6 else ""
            studenti.append(s)
            if s.strum1 and not s.doc1:
                note.append(f"{classe}ª {cognome}: manca il docente di {s.strum1} (1° strumento)")
                s.note.append(f"docente di {s.strum1} da indicare")
            if s.strum2 and not s.doc2:
                note.append(f"{classe}ª {cognome}: manca il docente di {s.strum2} (2° strumento)")
                s.note.append(f"docente di {s.strum2} da indicare")
    return studenti, note


def _docente_nella_riga(prima: str, pezzi: list[str], nomi_noti: set[str], famiglia: str) -> str:
    """Trova il nome del docente in righe come "PF 3 / Pacini C", "Saxofono Simonelli",
    "V.CELLO 1 - Di Bonaventura", "Pacini", "Percussione 2 / --------"."""
    # 1. un nome già visto nelle classi, scritto per intero
    for noto in sorted(nomi_noti, key=len, reverse=True):
        if norm(noto) in norm(prima):
            return noto
    # 2. il pezzo intero o le sue ultime parole, passati per gli alias ("Pacini" → "Pacini M")
    candidati: list[str] = []
    for pezzo in reversed([p for p in pezzi if p]):
        t = pezzo.replace("-", " ").strip()
        candidati.append(t)
        parole = t.split()
        for k in (2, 1):
            if len(parole) >= k:
                candidati.append(" ".join(parole[-k:]))
    for c in candidati:
        if set(c) <= set("-–— ") or not c.strip():
            continue
        if norm(c) in ALIAS_DOCENTI:
            return ALIAS_DOCENTI[norm(c)]
        d = docente(c)
        if d in nomi_noti:
            return d
    # 3. cattedra nuova: un nome proprio (una sola parola con la maiuscola) non è una sigla
    for c in candidati:
        if re.fullmatch(r"[A-Z][a-zà-ú']+", c.strip()) and norm(c) != norm(famiglia):
            return c.strip()
    return ""


def leggi_riassunti(doc, nomi_noti: set[str]) -> tuple[dict[str, dict], list[str]]:
    """Tabelle per strumento (pianoforte, fiati, chitarra, archi, canto).

    Da ogni riga tipo "PF 3 / Pacini C / 17 ore + 1 acc" ricava docente, ore e ore di
    accompagnamento. Ritorna {docente: {ore, accomp, cattedra, strumento}} e le segnalazioni.
    """
    info: dict[str, dict] = {}
    note: list[str] = []
    for ti in range(5, 10):
        t = doc.tables[ti]
        famiglia = celle(t.rows[0])[0] if t.rows else ""
        for row in t.rows[1:]:
            testo = celle(row)[0]
            if not testo or norm(testo).startswith("TOTALE") or "/" not in testo:
                continue
            pezzi = [p.strip() for p in testo.split("/")]
            # l'ultimo pezzo con "ora"/"h" sono le ore; il nome è il pezzo prima
            i_ore = next((i for i, p in enumerate(pezzi) if re.search(r"\d+\s*(ore|h)\b", p, re.I)), None)
            if i_ore is None:
                continue
            testo_ore = " ".join(pezzi[i_ore:])
            prima = " ".join(p for p in pezzi[:i_ore] if p)
            cattedra = pezzi[0].strip() if pezzi else famiglia
            d = _docente_nella_riga(prima, pezzi[:i_ore], nomi_noti, famiglia)
            m_ore = re.search(r"(\d+)\s*(?:ore|h)", testo_ore, re.I)
            m_acc = re.search(r"(\d+)\s*acc", testo_ore, re.I)
            if not d:
                note.append(f"cattedra senza titolare: «{testo.strip()}»"
                            + (f" ({m_ore.group(1)} ore" + (f" + {m_acc.group(1)} di accompagnamento" if m_acc else "") + ")" if m_ore else ""))
                continue
            info[d] = {"ore": int(m_ore.group(1)) if m_ore else None,
                       "accomp": int(m_acc.group(1)) if m_acc else 0,
                       "cattedra": cattedra, "famiglia": famiglia}
    return info, note


# ── Riconoscimento dei cognomi scritti nei gruppi ────────────────────────────

def pezzi_studente(cella: str) -> list[str]:
    """Divide una cella tipo "Gori C. PF/OB 3ALM / Gori Y. PF/Canto 5ALM" nei singoli ragazzi.

    La barra separa i ragazzi, ma è usata anche tra due strumenti ("PF/OB"): un pezzo che inizia
    con una sigla di strumento è la continuazione del ragazzo precedente.
    """
    grezzi = [p.strip() for p in re.split(r"[/,;]", cella) if p.strip()]
    out: list[str] = []
    for p in grezzi:
        primo = norm(p.split()[0]).strip(".")
        if out and (primo in SIGLE_STRUMENTO or re.fullmatch(r"\d?ALM", primo)):
            out[-1] += " " + p
        else:
            out.append(p)
    return out


def analizza_pezzo(pezzo: str) -> tuple[str, str, int | None]:
    """Da "Gori C. PF/OB 3ALM" ricava ("GORI", "C", 3): cognome, iniziale del nome, classe."""
    t = norm(pezzo)
    classe = None
    m = re.search(r"(\d)\s*ALM", t)
    if m:
        classe = int(m.group(1))
        t = t[:m.start()] + t[m.end():]
    parole = [w.strip(".,;()") for w in t.split() if w.strip(".,;()")]
    parole = [w for w in parole if w not in SIGLE_STRUMENTO and not re.fullmatch(r"\d+", w)]
    iniziale = ""
    if parole and len(parole[-1]) == 1:
        iniziale = parole.pop()
    return " ".join(parole), iniziale, classe


class Rubrica:
    """Trova uno studente a partire da come è scritto nei gruppi (cognome abbreviato, iniziale, classe)."""

    def __init__(self, studenti: list[Studente]):
        self.studenti = studenti

    def _candidati(self, cognome: str) -> list[Studente]:
        """Cognome esatto, oppure abbreviato: 'Ballati' → 'Ballati Niccolai', 'Padilla' → 'Roberto Padilla'."""
        parole = cognome.split()
        senza_apostrofi = cognome.replace("'", "")
        for prova in (
            lambda s: norm(s.cognome) == cognome,
            lambda s: norm(s.cognome).replace("'", "") == senza_apostrofi,
            lambda s: norm(s.cognome).split()[:len(parole)] == parole,     # prefisso
            lambda s: norm(s.cognome).split()[-len(parole):] == parole,    # suffisso
            lambda s: set(parole) <= set(norm(s.cognome).split()),         # parole contenute
        ):
            cand = [s for s in self.studenti if prova(s)]
            if cand:
                return cand
        return []

    def cerca(self, pezzo: str) -> tuple[Studente | None, str]:
        cognome, iniziale, classe = analizza_pezzo(pezzo)
        if not cognome:
            return None, "nessun cognome riconoscibile"
        cand = self._candidati(cognome)
        avviso = ""
        if not cand:  # refuso di battitura: cerca il cognome più somiglianteds
            tutti = {norm(s.cognome) for s in self.studenti}
            simili = difflib.get_close_matches(cognome, sorted(tutti), n=1, cutoff=0.85)
            if simili:
                cand = self._candidati(simili[0])
                avviso = f"scritto «{cognome.title()}», inteso come «{simili[0].title()}»"
        if classe is not None and len(cand) > 1:
            cand = [s for s in cand if s.classe == classe] or cand
        if iniziale and len(cand) > 1:
            cand = [s for s in cand if norm(s.nome).startswith(iniziale)] or cand
        if len(cand) == 1:
            s = cand[0]
            if classe is not None and s.classe != classe:
                avviso = (avviso + "; " if avviso else "") + f"scritto {classe}ALM ma in elenco è in {s.classe}ª"
            return s, avviso
        if not cand:
            return None, "non presente negli elenchi delle classi"
        return None, "più studenti corrispondono: " + ", ".join(f"{s.cognome} {s.nome} ({s.classe}ª)" for s in cand)

    def cerca_molti(self, pezzo: str) -> tuple[list[tuple[Studente, str]], list[str]]:
        """Come cerca(), ma accetta anche più ragazzi in un pezzo senza separatore
        ("Bonaguidi VL Greco VL"): divide sui cognomi che riconosce."""
        s, motivo = self.cerca(pezzo)
        if s is not None:
            return [(s, motivo)], []
        parole = [w for w in norm(pezzo).replace(".", " ").split()
                  if w not in SIGLE_STRUMENTO and not re.fullmatch(r"\d?ALM|\d+", w)]
        trovati: list[tuple[Studente, str]] = []
        i = 0
        while i < len(parole):
            preso = None
            for k in (2, 1):  # prova prima il cognome composto
                if i + k <= len(parole):
                    cand = self._candidati(" ".join(parole[i:i + k]))
                    if len(cand) == 1:
                        preso, i = (cand[0], i + k)
                        break
                    if len(cand) > 1 and k == 1:
                        break
            if preso is None:
                return [], [pezzo]
            trovati.append((preso, ""))
        return (trovati, []) if len(trovati) > 1 else ([], [pezzo])


# ── Gruppi di musica da camera ───────────────────────────────────────────────

@dataclass
class Gruppo:
    numero: int
    docente: str
    membri: list[Studente] = field(default_factory=list)
    note: list[str] = field(default_factory=list)


def leggi_gruppi(doc, rubrica: Rubrica) -> tuple[list[Gruppo], list[str]]:
    """La tabella delle FORMAZIONI: un gruppo per riga."""
    tabella = next((t for t in doc.tables
                    if t.rows and any(norm(x).startswith("FORMAZIONI") for x in celle(t.rows[0]))), None)
    if tabella is None:
        return [], ["tabella delle FORMAZIONI (gruppi di musica da camera) non trovata"]
    gruppi, note = [], []
    for row in tabella.rows:
        c = celle(row)
        if not c or not c[0].strip().isdigit():
            continue
        numero = int(c[0])
        i_conta = next((i for i, x in enumerate(c[1:], start=1) if x.strip().isdigit()), None)
        if i_conta is None:
            note.append(f"gruppo {numero}: manca il numero di partecipanti, riga saltata")
            continue
        dichiarati = int(c[i_conta])
        prof = docente(c[i_conta + 1]) if len(c) > i_conta + 1 else ""
        g = Gruppo(numero=numero, docente=prof)
        for cella in c[1:i_conta]:
            for pezzo in pezzi_studente(cella):
                trovati, dubbi = rubrica.cerca_molti(pezzo)
                for s, motivo in trovati:
                    g.membri.append(s)
                    if motivo:
                        note.append(f"gruppo {numero}: {s.cognome} — {motivo}")
                for x in dubbi:
                    _, motivo = rubrica.cerca(x)
                    g.note.append(f"«{x}»: {motivo}")
                    note.append(f"gruppo {numero}: «{x}» {motivo}")
        if not prof:
            note.append(f"gruppo {numero}: manca il docente")
        if len(g.membri) != dichiarati:
            note.append(f"gruppo {numero}: {len(g.membri)} ragazzi riconosciuti ma ne dichiara {dichiarati}")
        gruppi.append(g)
    return gruppi, note


# ── Laboratori di musica d'insieme (mattino) ─────────────────────────────────

@dataclass
class Laboratorio:
    nome: str
    classi: str
    docente: str
    aula: str
    giorno_ora: str
    studenti: str
    note: str = ""


def _giorno_ora(testo: str) -> str:
    m = re.search(r"(luned|marted|mercoled|gioved|venerd)\w*", testo, re.I)
    giorno = {"luned": "Lunedì", "marted": "Martedì", "mercoled": "Mercoledì",
              "gioved": "Giovedì", "venerd": "Venerdì"}.get(m.group(1).lower(), "") if m else ""
    ore = re.findall(r"(\d)\s*°?\s*(?:e\s*(\d)\s*°?)?\s*ora", testo, re.I)
    if ore:
        a, b = ore[0]
        return f"{giorno} {a}ª{f'-{b}ª' if b else ''} ora".strip()
    return giorno or testo.strip()[:40]


def _testa_laboratorio(testo: str) -> tuple[str, str, str]:
    """Da "ARCHI / Prof. Caselli / Aula M11, VC 1° piano" ricava (nome, docente, aula)."""
    pezzi = [p.strip() for p in re.split(r"\s{2,}|/|(?=Prof)|(?=Aula)", testo) if p and p.strip()]
    nome, prof, aula = "", "", ""
    for p in pezzi:
        if re.match(r"prof\b", p, re.I):
            prof = re.sub(r"^prof\.?\s*", "", p, flags=re.I).strip(" .,")
        elif re.match(r"aula\b", p, re.I):
            aula = p.strip()
        elif not nome:
            nome = p.strip()
    if not nome:
        nome = testo.strip()[:40]
    return nome, docente(prof), aula


def leggi_laboratori(doc, rubrica: Rubrica) -> tuple[list[Laboratorio], list[str]]:
    """La tabella MUSICA D'INSIEME: righe di intestazione (gruppi) alternate a righe di ragazzi.

    Le celle sono unite in modo irregolare: ogni gruppo prende le celle di ragazzi che ricadono
    nell'intervallo di colonne della sua intestazione.
    """
    tabella = next((t for t in doc.tables
                    if t.rows and any("INSIEME" in norm(x) for x in celle(t.rows[0]))), None)
    if tabella is None:
        return [], ["tabella MUSICA D'INSIEME (laboratori del mattino) non trovata"]
    lab, note = [], []
    righe = list(tabella.rows)
    quando, classi = "", ""
    i = 0
    while i < len(righe):
        colonne = celle_con_colonne(righe[i])
        testi = [t for _, _, t in colonne if t]
        unita = len(colonne) == 1
        if unita and testi and "INSIEME" in norm(testi[0]):
            quando = _giorno_ora(testi[0])
            m = re.search(r"class[ei]\s*([^–\-]*)", testi[0], re.I)
            grezzo = re.sub(r"\s*LMI.*", "", m.group(1)).strip(" :.") if m else ""
            numeri = re.findall(r"([1-5])", grezzo)
            classi = (f"{numeri[0]}ª" if len(numeri) == 1
                      else f"{numeri[0]}ª-{numeri[-1]}ª" if numeri else grezzo)
            i += 1
            continue
        intestazioni = [(c, n, t) for c, n, t in colonne if t and re.search(r"prof", t, re.I)]
        if not intestazioni:
            i += 1
            continue
        ragazzi_riga = celle_con_colonne(righe[i + 1]) if i + 1 < len(righe) else []
        for c, n, testo in intestazioni:
            nome, prof, aula = _testa_laboratorio(testo)
            pezzi: list[str] = []
            for c2, _, t2 in ragazzi_riga:
                if t2 and c <= c2 < c + n and not re.search(r"^\d+\s*ore", t2, re.I):
                    pezzi.append(t2)
            cognomi, dubbi = [], []
            for blocco in pezzi:
                for pezzo in pezzi_studente(re.sub(r"\d\s*ALM\s*:?", "", blocco)):
                    trovati, non_visti = rubrica.cerca_molti(pezzo)
                    for s, motivo in trovati:
                        cognomi.append(f"{s.cognome.title()} {s.classe}ª")
                        if motivo:
                            note.append(f"laboratorio «{nome}»: {s.cognome} — {motivo}")
                    dubbi += non_visti
            if dubbi:
                note.append(f"laboratorio «{nome}»: non riconosciuti {', '.join(dubbi)}")
            lab.append(Laboratorio(nome=f"LMI {nome}" if not norm(nome).startswith("LMI") else nome,
                                   classi=classi, docente=prof, aula=aula, giorno_ora=quando,
                                   studenti=", ".join(cognomi),
                                   note="da controllare: " + "; ".join(dubbi) if dubbi else ""))
        i += 2
    return lab, note


# ── Recapiti (comune, indirizzo, civico, km) ─────────────────────────────────

def spezza_civico(indirizzo: str) -> tuple[str, str]:
    t = re.sub(r"\s+", " ", (indirizzo or "").strip().strip(",.")).strip()
    m = re.match(r"^(.*?)[,\s]+(?:n\.?\s*|nr\.?\s*|n°\s*)?(\d+\s*[/\-]?\s*[A-Za-z0-9]{0,3})\s*$", t, re.I)
    if m and m.group(1):
        return m.group(1).strip(" ,.").title(), m.group(2).replace(" ", "").rstrip("/-").upper()
    return t.strip(" ,.").title(), ""


def leggi_recapiti(path: Path) -> dict[tuple[str, str], dict]:
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    testa = [norm(str(c.value)) for c in sh.row(0)]

    def col(prefisso: str) -> int | None:
        return next((i for i, h in enumerate(testa) if h.startswith(prefisso)), None)

    i_cog, i_nome = col("COGNOME"), col("NOME")
    i_km, i_com, i_ind = col("KM"), col("COMUNE"), col("INDIRIZZO")
    if i_cog is None or i_nome is None:
        raise SystemExit(f"{path.name}: non trovo le colonne Cognome e Nome")
    out: dict[tuple[str, str], dict] = {}
    for r in range(1, sh.nrows):
        cog, nome = norm(str(sh.cell_value(r, i_cog))), norm(str(sh.cell_value(r, i_nome)))
        if not cog:
            continue
        try:
            km = float(sh.cell_value(r, i_km)) if i_km is not None else None
        except (TypeError, ValueError):
            km = None
        via, civ = spezza_civico(str(sh.cell_value(r, i_ind))) if i_ind is not None else ("", "")
        out[(cog, nome)] = {"km": km, "indirizzo": via, "civico": civ,
                            "comune": str(sh.cell_value(r, i_com)).strip().title() if i_com is not None else ""}
    return out


def aggancia_recapiti(studenti: list[Studente], recapiti: dict) -> list[str]:
    note: list[str] = []
    senza: list[Studente] = []
    for s in studenti:
        cog, nome = s.chiave
        rec = recapiti.get((cog, nome))
        if rec is None:  # nome abbreviato o secondo nome in più
            primo = nome.split(" ")[0].rstrip(".")
            cand = [v for (c, n), v in recapiti.items() if c == cog and n.split(" ")[0] == primo]
            if len(cand) != 1:
                cand = [v for (c, n), v in recapiti.items() if c == cog]
            rec = cand[0] if len(cand) == 1 else None
        if rec is None:
            senza.append(s)
            continue
        s.km, s.comune, s.indirizzo, s.civico = rec["km"], rec["comune"], rec["indirizzo"], rec["civico"]
    if senza:
        per_classe = Counter(s.classe for s in senza)
        note.append(f"{len(senza)} studenti senza recapito (KM e indirizzo da inserire a mano), per classe: "
                    + ", ".join(f"{c}ª={n}" for c, n in sorted(per_classe.items())))
        for s in senza:
            s.note.append("recapito mancante")
        if len(senza) <= 12:
            note += [f"   {s.classe}ª {s.cognome} {s.nome}" for s in senza]
    return note


# ── Scrittura del file di input ──────────────────────────────────────────────

def scrivi(uscita: Path, studenti: list[Studente], docenti: dict[str, dict],
           gruppi: list[Gruppo], laboratori: list[Laboratorio]) -> None:
    righe_studenti = [{
        "classe": s.classe, "cognome": s.cognome, "nome": s.nome, "comune": s.comune,
        "indirizzo": s.indirizzo, "civico": s.civico, "km": s.km,
        "strum1": s.strum1, "doc1": s.doc1, "strum2": s.strum2, "doc2": s.doc2,
        "note": "; ".join(s.note),
    } for s in studenti]

    righe_docenti = []
    for nome in sorted(docenti):
        d = docenti[nome]
        note = DA_NOMINARE.get(nome, "")
        if d.get("senza_studenti"):
            note = (note + "; " if note else "") + (
                f"cattedra «{d.get('cattedra', '')}» nel riassunto, ma nessuno studente lo ha come docente")
        if d.get("ore"):
            note = (note + "; " if note else "") + f"nel Word: {d['ore']} ore" + (
                f" + {d['accomp']} di accompagnamento" if d.get("accomp") else "")
        righe_docenti.append({"nome": nome, "strumenti": ", ".join(sorted(d["strumenti"])),
                              "aula": AULE_PRECEDENTI.get(nome, ""), "ore_accomp": d.get("accomp", 0),
                              "note": note})

    wb = costruisci_workbook(righe_studenti, righe_docenti, esempi=False)

    ws = wb["Gruppi LMC"]
    for g in sorted(gruppi, key=lambda x: x.numero):
        cognomi = [m.cognome.title() if sum(1 for s in studenti if s.cognome == m.cognome) == 1
                   else f"{m.cognome.title()} {m.nome.title()}" for m in g.membri]
        ws.append([g.numero, g.docente] + cognomi + [""] * (5 - len(cognomi)) + ["; ".join(g.note)])

    ws = wb[FOGLIO_LMI]
    for l in laboratori:
        ws.append([l.nome, l.classi, l.docente, l.aula, l.giorno_ora, l.studenti, l.note])

    ws = wb[FOGLIO_PARAMETRI]
    for r in range(2, ws.max_row + 1):
        if "scuola" in str(ws.cell(row=r, column=1).value or "").lower():
            ws.cell(row=r, column=2).value = ws.cell(row=r, column=2).value or ""

    uscita.parent.mkdir(parents=True, exist_ok=True)
    wb.save(uscita)


def importa(docx_path: Path, recapiti_path: Path | None, uscita: Path) -> None:
    doc = docx.Document(docx_path)
    print(f"Leggo {docx_path.name}")

    studenti, note_classi = leggi_classi(doc)
    nomi_noti = {d for s in studenti for d in (s.doc1, s.doc2) if d}
    riassunti, note_riassunti = leggi_riassunti(doc, nomi_noti)
    rubrica = Rubrica(studenti)
    gruppi, note_gruppi = leggi_gruppi(doc, rubrica)
    laboratori, note_lab = leggi_laboratori(doc, rubrica)

    note_recapiti: list[str] = []
    if recapiti_path and recapiti_path.exists():
        print(f"Leggo {recapiti_path.name}")
        note_recapiti = aggancia_recapiti(studenti, leggi_recapiti(recapiti_path))

    # elenco docenti: da chi insegna, da chi tiene un gruppo, da chi compare nei riassunti
    docenti: dict[str, dict] = defaultdict(lambda: {"strumenti": set(), "accomp": 0, "ore": None})
    for s in studenti:
        for d, strum in ((s.doc1, s.strum1), (s.doc2, s.strum2)):
            if d:
                docenti[d]["strumenti"].add(strum)
    for g in gruppi:
        if g.docente:
            docenti[g.docente]["strumenti"].add("MUSICA DA CAMERA")
    for l in laboratori:
        if l.docente:
            docenti[l.docente]["strumenti"].add("MUSICA D'INSIEME")
    for nome, info in riassunti.items():
        docenti[nome]["accomp"] = info["accomp"]
        docenti[nome]["ore"] = info["ore"]
        docenti[nome]["cattedra"] = info["cattedra"]
        if not docenti[nome]["strumenti"] and info["famiglia"]:
            # docente senza studenti negli elenchi: almeno la famiglia di strumento dalla cattedra
            docenti[nome]["strumenti"].add(strumento(info["famiglia"]))
            docenti[nome]["senza_studenti"] = True

    scrivi(uscita, studenti, docenti, gruppi, laboratori)

    # ── riepilogo ──
    print(f"\nCreato: {uscita}")
    print(f"  studenti: {len(studenti)} — " + ", ".join(
        f"{c}ª={sum(1 for s in studenti if s.classe == c)}" for c in range(1, 6)))
    print(f"  docenti: {len(docenti)}")
    print(f"  gruppi di musica da camera: {len(gruppi)}, con {sum(len(g.membri) for g in gruppi)} ragazzi")
    print(f"  laboratori del mattino (LMI): {len(laboratori)}")
    con_km = sum(1 for s in studenti if s.km is not None)
    con_ind = sum(1 for s in studenti if s.indirizzo)
    print(f"  recapiti agganciati: {con_km} con i km, {con_ind} con l'indirizzo")
    accomp = {n: d["accomp"] for n, d in docenti.items() if d.get("accomp")}
    if accomp:
        print("  ore di accompagnamento: " + ", ".join(f"{n} {v}" for n, v in accomp.items()))
    # confronto tra le ore dichiarate nel riassunto e quelle che risultano dagli elenchi delle classi
    ore_reali: Counter = Counter()
    for s in studenti:
        h1 = 2 if s.classe in (1, 2, 5) else 1
        if s.doc1:
            ore_reali[s.doc1] += h1
        if s.doc2 and s.classe != 5:
            ore_reali[s.doc2] += 1
    differenze = []
    for nome, info in sorted(riassunti.items()):
        atteso = info.get("ore")
        if atteso is None:
            continue
        vere = ore_reali.get(nome, 0)
        if vere != atteso:
            differenze.append(f"{nome}: nel riassunto {atteso} ore di strumento, negli elenchi delle classi {vere}"
                              f" ({'+' if vere > atteso else ''}{vere - atteso})")

    triennio = [s for s in studenti if s.classe in (3, 4, 5)]
    in_gruppo = {m.chiave for g in gruppi for m in g.membri}
    senza = [s for s in triennio if s.chiave not in in_gruppo]
    doppi = Counter(m.chiave for g in gruppi for m in g.membri)

    tutte = ([f"[classi] {x}" for x in note_classi] + [f"[riassunti] {x}" for x in note_riassunti]
             + [f"[gruppi] {x}" for x in note_gruppi] + [f"[LMI] {x}" for x in note_lab]
             + [f"[recapiti] {x}" for x in note_recapiti]
             + [f"[gruppi] {s.classe}ª {s.cognome} {s.nome} non è in nessun gruppo" for s in senza]
             + [f"[gruppi] {c[0]} {c[1]} è in {n} gruppi" for c, n in doppi.items() if n > 1]
             + [f"[ore] {x}" for x in differenze])
    if tutte:
        print(f"\nDa controllare ({len(tutte)}):")
        for x in tutte:
            print("   -", x)
    else:
        print("\nNessuna segnalazione: tutto riconosciuto.")


def main(argv: list[str]) -> None:
    def piu_recente(schema: str) -> Path | None:
        trovati = sorted(ROOT.glob(schema), key=lambda p: p.stat().st_mtime, reverse=True)
        return trovati[0] if trovati else None

    docx_path = Path(argv[0]) if len(argv) > 0 else piu_recente("classi e gruppi*.docx")
    recapiti = Path(argv[1]) if len(argv) > 1 else piu_recente("Recapiti*.xls*")
    uscita = Path(argv[2]) if len(argv) > 2 else USCITA_DEFAULT
    if docx_path is None or not docx_path.exists():
        raise SystemExit("Non trovo il file Word 'classi e gruppi ....docx'. Passalo come primo argomento.")
    importa(docx_path, recapiti, uscita)


if __name__ == "__main__":
    main(sys.argv[1:])
