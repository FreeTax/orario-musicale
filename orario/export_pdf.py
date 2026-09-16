"""Esporta l'orario in PDF (reportlab/platypus): settimanale, per docente, per studente."""

from __future__ import annotations

from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

from .costanti import GIORNI, GIORNI_LUNGHI, N_GIORNI, N_ORE, ORE, ORE_FINE, ORE_LABEL
from .export_comune import (
    LEGENDA, LEGENDA_CASELLE,
    Cella, cognomi_doppi, etichetta_fascia, griglia_docente, griglia_studente,
    studenti_ordinati, titolo_orario,
)
from .modello import Orario

GRIGIO = colors.HexColor("#E7E6E6")
AZZURRO = colors.HexColor("#DDEBF7")
BORDO = colors.HexColor("#888888")
MARGINE = 10 * mm



# ── Stili ────────────────────────────────────────────────────────────────────

def _stile(nome: str, size: float, bold: bool = False, align=TA_CENTER, leading: float | None = None) -> ParagraphStyle:
    return ParagraphStyle(
        nome, fontName="Helvetica-Bold" if bold else "Helvetica", fontSize=size,
        leading=leading or size * 1.15, alignment=align, splitLongWords=1,
    )


def _par(cella: Cella, stile: ParagraphStyle) -> Paragraph:
    """Paragraph con una riga per ogni riga della cella, con grassetto/corsivo."""
    pezzi = []
    for r in cella.righe:
        t = escape(r.testo)
        if r.grassetto:
            t = f"<b>{t}</b>"
        if r.corsivo:
            t = f"<i>{t}</i>"
        pezzi.append(t)
    return Paragraph("<br/>".join(pezzi), stile)


def _testo(testo: str, stile: ParagraphStyle) -> Paragraph:
    return Paragraph(escape(testo), stile)


def _stile_tabella(n_righe_int: int, n_righe: int, n_col: int, grigie: list[tuple[int, int]]) -> TableStyle:
    comandi = [
        ("GRID", (0, 0), (-1, -1), 0.5, BORDO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, n_righe_int - 1), AZZURRO),
        ("BACKGROUND", (0, n_righe_int), (0, n_righe - 1), AZZURRO),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    for c, r in grigie:
        comandi.append(("BACKGROUND", (c, r), (c, r), GRIGIO))
    return TableStyle(comandi)


# ── Orario settimanale ───────────────────────────────────────────────────────

def _dimensiona(testo: str, larghezza: float, size: float, font: str = "Helvetica-Bold", minimo: float = 4.5) -> float:
    """Riduce il corpo del carattere finché il testo sta su una riga nella larghezza data."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    w = stringWidth(testo, font, size)
    if w <= larghezza:
        return size
    return max(minimo, size * larghezza / w)


def _tabella_giorno(orario: Orario, g: int, doppi: set[str], mappa, larghezza_utile: float, altezza_utile: float) -> Table:
    dati = orario.dati
    n = len(dati.docenti)
    size = 7.5 if n <= 14 else 7 if n <= 18 else 6.5 if n <= 21 else 6
    st_int_n = _stile("intn", size)
    st_cel = _stile("cel", size)
    st_fascia = _stile("fascia", size, bold=True)

    larg_fascia = 42
    larg_col = (larghezza_utile - larg_fascia) / max(n, 1)
    interno = larg_col - 4  # padding sinistro + destro

    def nome(d) -> Paragraph:
        # cognomi composti possono andare a capo tra le parole; ogni parola deve stare in una riga
        parola = max(d.nome.split() or [d.nome], key=len)
        return _testo(d.nome, _stile("nome", _dimensiona(parola, interno, size), bold=True))

    def aula(d) -> Paragraph:
        testo = d.aula(g)
        if not testo:
            return _testo("-", st_int_n)
        return Paragraph(f"AULA<br/>{escape(testo)}", _stile("aula", _dimensiona(testo, interno, size), bold=True))

    righe = [
        [_testo(GIORNI_LUNGHI[g], _stile("int", size, bold=True))] + [nome(d) for d in dati.docenti],
        [_testo("Aula", st_int_n)] + [aula(d) for d in dati.docenti],
        [_testo("Strumento", st_int_n)] + [_testo(d.strumenti.title() or "-", st_int_n) for d in dati.docenti],
    ]
    grigie: list[tuple[int, int]] = []
    griglie = [griglia_docente(orario, d.nome, doppi, mappa) for d in dati.docenti]
    for o in range(N_ORE):
        riga = [_testo(etichetta_fascia(o), st_fascia)]
        for j, gr in enumerate(griglie):
            c = gr[o][g]
            riga.append(_par(c, st_cel) if not c.vuota else "")
            if c.vuota and c.non_disponibile:
                grigie.append((j + 1, 3 + o))
        righe.append(riga)
    larghezze = [larg_fascia] + [larg_col] * n
    stile = _stile_tabella(3, len(righe), n + 1, grigie)

    # prima passata: altezze naturali; poi le 4 fasce si allargano fino a riempire la pagina
    prova = Table(righe, colWidths=larghezze)
    prova.setStyle(stile)
    prova.wrap(larghezza_utile, 10_000)
    naturali = list(prova._rowHeights)
    altezza_intestazione = sum(naturali[:3])
    disponibile = altezza_utile - altezza_intestazione
    altezza_fascia = max(max(naturali[3:]), disponibile / N_ORE) if disponibile > 0 else max(naturali[3:])
    t = Table(righe, colWidths=larghezze, rowHeights=[None] * 3 + [altezza_fascia] * N_ORE)
    t.setStyle(stile)
    return t


def _pagina_lmi(orario: Orario, larghezza_utile: float) -> list:
    st_tit = _stile("tit", 13, bold=True)
    st_int = _stile("int", 8, bold=True)
    st_cel = _stile("cel", 8, align=0)
    colonne = ["Laboratorio", "Classi", "Docente", "Aula", "Giorno e ora", "Studenti"]
    quote = [0.16, 0.07, 0.12, 0.15, 0.14, 0.36]
    righe = [[_testo(c, st_int) for c in colonne]]
    for l in orario.dati.lmi:
        righe.append([_testo(v or "-", st_cel) for v in (l.nome, l.classi, l.docente, l.aula, l.giorno_ora, l.studenti)])
    t = Table(righe, colWidths=[q * larghezza_utile for q in quote], repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, BORDO),
        ("BACKGROUND", (0, 0), (-1, 0), AZZURRO),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return [PageBreak(), _testo(titolo_orario(orario.dati), st_tit), Spacer(1, 4),
            _testo("LABORATORI DI MUSICA D'INSIEME (ORARIO DEL MATTINO)", _stile("sot", 11, bold=True)),
            Spacer(1, 8), t]


def esporta_pdf_settimanale(orario: Orario, percorso: Path) -> Path:
    percorso = Path(percorso)
    percorso.parent.mkdir(parents=True, exist_ok=True)
    pagina = landscape(A4)
    doc = SimpleDocTemplate(str(percorso), pagesize=pagina, leftMargin=MARGINE, rightMargin=MARGINE,
                            topMargin=MARGINE, bottomMargin=MARGINE, title=titolo_orario(orario.dati))
    larghezza_utile = pagina[0] - 2 * MARGINE
    doppi = cognomi_doppi(orario.dati)
    mappa = orario.per_docente_fascia()
    st_tit = _stile("tit", 13, bold=True)
    st_giorno = _stile("giorno", 11, bold=True)
    st_leg = _stile("leg", 7)
    # spazio verticale per la tabella: pagina - margini - titolo/giorno/legenda
    altezza_utile = pagina[1] - 2 * MARGINE - (13 * 1.15 + 4 + 11 * 1.15 + 6) - (7 * 1.15 * 2 + 8) - 6
    flussi: list = []
    for g in range(N_GIORNI):
        if g:
            flussi.append(PageBreak())
        flussi += [
            _testo(titolo_orario(orario.dati), st_tit), Spacer(1, 4),
            _testo(GIORNI_LUNGHI[g].upper(), st_giorno), Spacer(1, 6),
            _tabella_giorno(orario, g, doppi, mappa, larghezza_utile, altezza_utile),
            Spacer(1, 8), _testo(LEGENDA + " — " + LEGENDA_CASELLE, st_leg),
        ]
    if orario.dati.lmi:
        flussi += _pagina_lmi(orario, larghezza_utile)
    doc.build(flussi)
    return percorso


# ── Griglia 4 × 5 (docenti e studenti) ───────────────────────────────────────

def _griglia_settimana(griglia: list[list[Cella]], larghezza: float, size: float, altezza_riga: float) -> Table:
    st_int = _stile("int", size, bold=True)
    st_cel = _stile("cel", size)
    larg_fascia = 58
    larg_col = (larghezza - larg_fascia) / N_GIORNI
    righe = [[_testo("", st_int)] + [_testo(gg, st_int) for gg in GIORNI_LUNGHI]]
    grigie: list[tuple[int, int]] = []
    for o in range(N_ORE):
        riga = [_testo(etichetta_fascia(o), st_int)]
        for g in range(N_GIORNI):
            c = griglia[o][g]
            riga.append(_par(c, st_cel) if not c.vuota else "")
            if c.vuota and c.non_disponibile:
                grigie.append((g + 1, 1 + o))
        righe.append(riga)
    t = Table(righe, colWidths=[larg_fascia] + [larg_col] * N_GIORNI, rowHeights=[None] + [altezza_riga] * N_ORE)
    t.setStyle(_stile_tabella(1, len(righe), N_GIORNI + 1, grigie))
    return t


# ── Orario per docente ───────────────────────────────────────────────────────

def esporta_pdf_docenti(orario: Orario, percorso: Path) -> Path:
    percorso = Path(percorso)
    percorso.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(percorso), pagesize=A4, leftMargin=MARGINE * 1.5, rightMargin=MARGINE * 1.5,
                            topMargin=MARGINE * 1.5, bottomMargin=MARGINE * 1.5,
                            title=titolo_orario(orario.dati) + " - docenti")
    larghezza_utile = A4[0] - 3 * MARGINE
    doppi = cognomi_doppi(orario.dati)
    mappa = orario.per_docente_fascia()
    st_tit = _stile("tit", 12, bold=True)
    st_nome = _stile("nome", 16, bold=True)
    st_info = _stile("info", 10)
    st_leg = _stile("leg", 8)
    flussi: list = []
    for i, d in enumerate(orario.dati.docenti):
        if i:
            flussi.append(PageBreak())
        lezioni = orario.lezioni_docente(d.nome)
        if d.aula_unica:
            aule = f"Aula: {d.aula() or '-'}"
        else:
            aule = "Aule: " + ", ".join(f"{GIORNI[i]} {a or '-'}" for i, a in enumerate(d.aule))
        info = [f"Strumento: {d.strumenti.title() or '-'}", aule,
                f"Ore in orario: {len(lezioni)} (disponibilità: {len(d.disponibilita)} fasce)"]
        flussi += [
            _testo(titolo_orario(orario.dati), st_tit), Spacer(1, 10),
            _testo(d.nome, st_nome), Spacer(1, 4),
            _testo("   |   ".join(info), st_info), Spacer(1, 14),
            _griglia_settimana(griglia_docente(orario, d.nome, doppi, mappa), larghezza_utile, 9, 92),
            Spacer(1, 10), _testo(LEGENDA + " — " + LEGENDA_CASELLE, st_leg),
        ]
    doc.build(flussi)
    return percorso


# ── Orario per studente ──────────────────────────────────────────────────────

def esporta_pdf_studenti(orario: Orario, percorso: Path) -> Path:
    percorso = Path(percorso)
    percorso.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(percorso), pagesize=A4, leftMargin=MARGINE * 1.5, rightMargin=MARGINE * 1.5,
                            topMargin=MARGINE * 1.5, bottomMargin=MARGINE * 1.5,
                            title=titolo_orario(orario.dati) + " - studenti")
    larghezza_utile = A4[0] - 3 * MARGINE
    doppi = cognomi_doppi(orario.dati)
    st_tit = _stile("tit", 10, bold=True)
    st_nome = _stile("nome", 14, bold=True)
    st_info = _stile("info", 9.5)
    st_leg = _stile("leg", 7.5)
    flussi: list = []
    studenti = studenti_ordinati(orario.dati)
    for i, s in enumerate(studenti):
        blocco = [
            _testo(titolo_orario(orario.dati), st_tit), Spacer(1, 6),
            _testo(f"{s.cognome.title()} {s.nome.title()}", st_nome), Spacer(1, 3),
            _testo(f"Classe {s.classe}ª", st_info), Spacer(1, 8),
            _griglia_settimana(griglia_studente(orario, s.id, doppi), larghezza_utile, 8, 52),
            Spacer(1, 6), _testo(LEGENDA + " — " + LEGENDA_CASELLE, st_leg),
        ]
        flussi.append(KeepTogether(blocco))
        if i % 2 == 1:
            flussi.append(PageBreak())
        else:
            flussi.append(Spacer(1, 26))
    if flussi and isinstance(flussi[-1], Spacer):
        flussi.pop()
    doc.build(flussi)
    return percorso


__all__ = ["esporta_pdf_settimanale", "esporta_pdf_docenti", "esporta_pdf_studenti"]
