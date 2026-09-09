"""Helper condivisi da export_excel ed export_pdf: testo delle celle, etichette, riepiloghi."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from .costanti import (
    ETICHETTA_ACCOMP, GIORNI, MAX_GRUPPO_LMC, N_ORE, ORE, ORE_FINE, ORE_LABEL,
    TIPO_ACCOMP, TIPO_LMC, TIPO_STRUM1, TIPO_STRUM2,
)
from .modello import DatiInput, Lezione, Orario, Studente


# ── Righe di testo con stile ─────────────────────────────────────────────────

@dataclass
class Riga:
    testo: str
    grassetto: bool = False
    corsivo: bool = False


@dataclass
class Cella:
    """Contenuto di una casella dell'orario (docente × fascia)."""
    righe: list[Riga] = field(default_factory=list)
    non_disponibile: bool = False  # il docente non ha X/A in questa fascia
    lezione: Lezione | None = None

    @property
    def vuota(self) -> bool:
        return not self.righe

    @property
    def testo(self) -> str:
        return "\n".join(r.testo for r in self.righe)

    @property
    def grassetto(self) -> bool:
        return bool(self.righe) and all(r.grassetto for r in self.righe)

    @property
    def corsivo(self) -> bool:
        return bool(self.righe) and all(r.corsivo for r in self.righe)


# ── Etichette ────────────────────────────────────────────────────────────────

def cognomi_doppi(dati: DatiInput) -> set[str]:
    """Cognomi (maiuscolo) portati da più di uno studente, in qualunque classe."""
    c = Counter(s.cognome.upper().strip() for s in dati.studenti)
    return {k for k, v in c.items() if v > 1}


def etichetta_studente(s: Studente, doppi: set[str]) -> str:
    return s.etichetta(con_iniziale=s.cognome.upper().strip() in doppi)


def etichetta_id(dati: DatiInput, id_: str, doppi: set[str]) -> str:
    s = dati.studente(id_)
    return etichetta_studente(s, doppi) if s else id_.title()


def etichetta_fascia(ora: int) -> str:
    """'7ª 13:30-14:30'."""
    return f"{ORE_LABEL[ora]} {ORE[ora]}-{ORE_FINE[ora]}"


def etichetta_fascia_breve(ora: int) -> str:
    return f"{ORE_LABEL[ora]} {ORE[ora]}"


def anno_scolastico(percorso: Path | None) -> str:
    """Estrae '2026-27' dal nome del file di input, se c'è; altrimenti ''."""
    if not percorso:
        return ""
    m = re.search(r"(20\d{2})\s*[-_/]\s*(20)?(\d{2})", Path(percorso).stem)
    if not m:
        return ""
    return f"{m.group(1)}-{m.group(3)}"


def titolo_orario(dati: DatiInput) -> str:
    anno = anno_scolastico(dati.percorso)
    return "LICEO MUSICALE: ORARIO POMERIDIANO" + (f" {anno}" if anno else "")


def descrizione_docente(dati: DatiInput, nome: str) -> str:
    """'Bini (aula 5ALM)' oppure 'Bini' se l'aula manca."""
    d = dati.docente(nome)
    if d and d.aula:
        return f"{nome} (aula {d.aula})"
    return nome


# ── Contenuto delle celle ────────────────────────────────────────────────────

def cella_lezione(orario: Orario, lez: Lezione, doppi: set[str], molti_lmc: int = MAX_GRUPPO_LMC) -> Cella:
    """Costruisce la cella per una lezione (senza informazioni sulla disponibilità)."""
    dati = orario.dati
    righe: list[Riga] = []
    if lez.tipo in (TIPO_STRUM1, TIPO_STRUM2):
        b, i = lez.tipo == TIPO_STRUM1, lez.tipo == TIPO_STRUM2
        for id_ in lez.studenti:
            righe.append(Riga(etichetta_id(dati, id_, doppi), b, i))
        if lez.due_ore:
            righe.append(Riga("(2ª ora)", b, i))
    elif lez.tipo == TIPO_LMC:
        righe.append(Riga("LMC"))
        nomi = [etichetta_id(dati, id_, doppi) for id_ in lez.studenti]
        if len(nomi) > molti_lmc:
            righe.append(Riga(", ".join(nomi)))
        else:
            righe.extend(Riga(n) for n in nomi)
    elif lez.tipo == TIPO_ACCOMP:
        righe.append(Riga(ETICHETTA_ACCOMP))
    else:
        righe.append(Riga(lez.tipo))
    return Cella(righe=righe, lezione=lez)


def cella(orario: Orario, docente: str, fascia: int, doppi: set[str],
          mappa: dict[tuple[str, int], Lezione] | None = None) -> Cella:
    """Cella per (docente, fascia): lezione, oppure vuota/grigia se non disponibile."""
    if mappa is None:
        mappa = orario.per_docente_fascia()
    lez = mappa.get((docente, fascia))
    if lez is not None:
        return cella_lezione(orario, lez, doppi)
    d = orario.dati.docente(docente)
    return Cella(non_disponibile=bool(d) and not d.disponibile(fascia))


def griglia_docente(orario: Orario, docente: str, doppi: set[str],
                    mappa: dict[tuple[str, int], Lezione] | None = None) -> list[list[Cella]]:
    """Matrice 4 ore × 5 giorni delle celle di un docente."""
    if mappa is None:
        mappa = orario.per_docente_fascia()
    return [[cella(orario, docente, g * N_ORE + o, doppi, mappa) for g in range(len(GIORNI))]
            for o in range(N_ORE)]


# ── Vista studente ───────────────────────────────────────────────────────────

def tipo_lezione_breve(lez: Lezione) -> str:
    return {TIPO_STRUM1: "1° str.", TIPO_STRUM2: "2° str.", TIPO_LMC: "LMC", TIPO_ACCOMP: "Accomp."}.get(lez.tipo, lez.tipo)


def tipo_lezione_lungo(lez: Lezione) -> str:
    return {TIPO_STRUM1: "1° strumento", TIPO_STRUM2: "2° strumento", TIPO_LMC: "Musica da camera",
            TIPO_ACCOMP: "Accompagnamento"}.get(lez.tipo, lez.tipo)


def lezione_per_studente_breve(orario: Orario, lez: Lezione) -> str:
    """'13:30 1° str. Bini' (per il foglio Studenti dell'Excel)."""
    testo = f"{ORE[lez.ora]} {tipo_lezione_breve(lez)} {lez.docente}"
    if lez.due_ore:
        testo += " (2ª ora)"
    return testo


def lezione_per_studente(orario: Orario, lez: Lezione, id_: str, doppi: set[str]) -> Cella:
    """Cella della griglia personale dello studente: '1° strumento – Bini (aula M1)' ecc."""
    dati = orario.dati
    b, i = lez.tipo == TIPO_STRUM1, lez.tipo == TIPO_STRUM2
    righe = [Riga(tipo_lezione_lungo(lez), b, i), Riga(descrizione_docente(dati, lez.docente), b, i)]
    if lez.tipo == TIPO_LMC:
        altri = [etichetta_id(dati, x, doppi) for x in lez.studenti if x != id_]
        if altri:
            righe.append(Riga("con " + ", ".join(altri)))
    if lez.due_ore:
        righe.append(Riga("(2ª ora)", b, i))
    return Cella(righe=righe, lezione=lez)


def griglia_studente(orario: Orario, id_: str, doppi: set[str]) -> list[list[Cella]]:
    """Matrice 4 ore × 5 giorni delle lezioni di uno studente."""
    griglia = [[Cella() for _ in GIORNI] for _ in range(N_ORE)]
    for lez in orario.lezioni_di(id_):
        griglia[lez.ora][lez.giorno] = lezione_per_studente(orario, lez, id_, doppi)
    return griglia


def giorni_rientro(orario: Orario, id_: str) -> list[str]:
    return [GIORNI[g] for g in sorted({l.giorno for l in orario.lezioni_di(id_)})]


def studenti_ordinati(dati: DatiInput) -> list[Studente]:
    return sorted(dati.studenti, key=lambda s: (s.classe, s.cognome.upper(), s.nome.upper()))


# ── Riepilogo per il foglio Controlli ────────────────────────────────────────

def riepilogo(orario: Orario) -> list[tuple[str, object]]:
    """Coppie (voce, valore) con i numeri principali dell'orario."""
    dati = orario.dati
    p = dati.parametri
    rientri = Counter(orario.rientri_di(s.id) for s in dati.studenti)
    buchi_tot = sum(orario.buchi_di(s.id) for s in dati.studenti)
    lontani = {s.id for s in dati.studenti if s.km is not None and s.km > p.soglia_km_vicino}
    ultime_lontani = sum(1 for l in orario.lezioni if l.ora == N_ORE - 1
                         and any(x in lontani for x in l.studenti))
    oltre_max = 0
    for s in dati.studenti:
        limite = p.max_rientri_vicini if (s.km is not None and s.km <= p.soglia_km_vicino) else p.max_rientri
        if orario.rientri_di(s.id) > limite:
            oltre_max += 1
    n_lmc = sum(1 for l in orario.lezioni if l.tipo == TIPO_LMC)
    n_acc = sum(1 for l in orario.lezioni if l.tipo == TIPO_ACCOMP)
    righe: list[tuple[str, object]] = [
        ("Studenti", len(dati.studenti)),
        ("Docenti", len(dati.docenti)),
        ("Lezioni in orario", len(orario.lezioni)),
        ("  di cui musica da camera (LMC)", n_lmc),
        ("  di cui accompagnamento", n_acc),
        ("Stato del calcolo", orario.stato or "-"),
        ("Tempo di calcolo (secondi)", round(orario.secondi, 1)),
        ("Totale buchi (ore libere tra due lezioni)", buchi_tot),
        ("Studenti senza lezioni", rientri.get(0, 0)),
        ("Studenti con 1 rientro", rientri.get(1, 0)),
        ("Studenti con 2 rientri", rientri.get(2, 0)),
        ("Studenti con 3 rientri", rientri.get(3, 0)),
        ("Studenti con 4 o più rientri", sum(v for k, v in rientri.items() if k >= 4)),
        (f"Studenti oltre il massimo rientri ({p.max_rientri}, {p.max_rientri_vicini} entro {p.soglia_km_vicino:g} km)", oltre_max),
        (f"Lezioni alla {ORE_LABEL[-1]} ora di studenti lontani (oltre {p.soglia_km_vicino:g} km)", ultime_lontani),
        ("Avvisi del programma", len(orario.avvisi)),
    ]
    return righe
