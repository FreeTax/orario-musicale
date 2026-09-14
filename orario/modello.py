"""Strutture dati condivise tra lettura, controlli, motore ed export."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .costanti import ORE_PER_CLASSE, TIPO_ACCOMP, TIPO_LMC, giorno_ora


# ── Input ────────────────────────────────────────────────────────────────────

@dataclass
class Trasporto:
    """Ritorno a casa con i mezzi pubblici, per ciascuna delle 4 fasce (fine lezione → arrivo a casa)."""
    minuti: list[int | None]  # per fascia 0..3: minuti da fine lezione ad arrivo a casa; None = nessun mezzo
    arrivi: list[str] = field(default_factory=lambda: ["", "", "", ""])  # ora di arrivo a casa, testo
    mezzi: list[str] = field(default_factory=lambda: ["", "", "", ""])  # descrizione del percorso
    lat: float | None = None
    lon: float | None = None
    indirizzo_usato: str = ""
    esito: str = ""  # "OK", "indirizzo non trovato", "nessun percorso", ...
    aggiornato: str = ""

    @property
    def valido(self) -> bool:
        return any(m is not None for m in self.minuti)

    @property
    def minuti_min(self) -> int | None:
        validi = [m for m in self.minuti if m is not None]
        return min(validi) if validi else None


@dataclass
class Studente:
    classe: int
    cognome: str
    nome: str
    km: float | None
    strum1: str
    doc1: str
    strum2: str
    doc2: str
    primo_attaccato: bool | None = None  # SI = le 2 ore di 1° strumento attaccate; NO o vuoto = giorni diversi
    giorno_unico: bool = False
    giorni_non_disp: set[int] = field(default_factory=set)  # indici 0..4
    fasce_non_disp: set[int] = field(default_factory=set)   # singole ore occupate, indici 0..19
    note: str = ""
    riga: int = 0  # riga nel foglio Excel (per i messaggi)
    comune: str = ""
    indirizzo: str = ""
    civico: str = ""
    trasporto: Trasporto | None = None  # riempito da "Aggiorna trasporti" (foglio Trasporti)
    minuti_manuali: float | None = None  # colonna «Minuti per tornare a casa» scritta a mano: vince su tutto

    @property
    def minuti_ritorno(self) -> float:
        """Quanto ci mette a tornare a casa, in minuti: dai mezzi se ci sono, altrimenti stimati dai km."""
        if self.minuti_manuali is not None:
            return float(self.minuti_manuali)
        if self.trasporto is not None and self.trasporto.minuti_min is not None:
            return float(self.trasporto.minuti_min)
        if self.km is not None:
            return 10.0 + 2.5 * self.km
        return 0.0

    @property
    def ore_consecutive(self) -> bool:
        """Le 2 ore di 1° strumento devono essere una di seguito all'altra."""
        return self.primo_attaccato is True

    @property
    def primo_separato(self) -> bool:
        """Le 2 ore di 1° strumento devono stare in giorni diversi: è il comportamento normale.

        Solo un SI esplicito le attacca; la casella vuota vale come NO."""
        return self.primo_attaccato is not True

    @property
    def indirizzo_completo(self) -> str:
        via = " ".join(x for x in (self.indirizzo.strip(), self.civico.strip()) if x)
        return ", ".join(x for x in (via, self.comune.strip()) if x)

    @property
    def id(self) -> str:
        """Identificativo unico: 'COGNOME NOME'."""
        return f"{self.cognome.upper()} {self.nome.upper()}".strip()

    @property
    def ore(self) -> tuple[int, int, int]:
        """(ore 1° strumento, ore 2° strumento, ore LMC) per la sua classe."""
        return ORE_PER_CLASSE[self.classe]

    def libero(self, f: int) -> bool:
        """Il ragazzo può venire in quella fascia: né giorno vietato né impegno personale."""
        return giorno_ora(f)[0] not in self.giorni_non_disp and f not in self.fasce_non_disp

    @property
    def fasce_vietate(self) -> set[int]:
        """Tutte le ore in cui non può esserci: giorni interi più singoli impegni."""
        return {f for f in range(20) if not self.libero(f)}

    def etichetta(self, con_iniziale: bool = False) -> str:
        """Testo da scrivere nell'orario: cognome (+ iniziale se richiesto) + classe."""
        base = self.cognome.title()
        if con_iniziale and self.nome:
            base += f" {self.nome[0].upper()}."
        return f"{base} {self.classe}ª"


@dataclass
class Docente:
    nome: str
    strumenti: str = ""
    aule: list[str] = field(default_factory=lambda: ["", "", "", "", ""])  # una per giorno, lun→ven
    ore_accomp: int = 0
    disponibilita: dict[int, str] = field(default_factory=dict)  # fascia → "X" oppure "A"
    note: str = ""
    riga: int = 0

    def aula(self, giorno: int | None = None) -> str:
        """L'aula di quel giorno. Senza giorno: l'unica aula, se è sempre la stessa, altrimenti "".."""
        if giorno is not None:
            return self.aule[giorno] if giorno < len(self.aule) else ""
        distinte = {a for a in self.aule if a}
        return distinte.pop() if len(distinte) == 1 else ""

    @property
    def aula_unica(self) -> bool:
        return len({a for a in self.aule if a}) <= 1

    def disponibile(self, f: int) -> bool:
        return f in self.disponibilita

    @property
    def fasce_x(self) -> list[int]:
        return sorted(f for f, v in self.disponibilita.items() if v == "X")

    @property
    def fasce_a(self) -> list[int]:
        return sorted(f for f, v in self.disponibilita.items() if v == "A")


@dataclass
class GruppoLMC:
    numero: int
    docente: str
    studenti: list[str]  # id studente ("COGNOME NOME"), già risolti dai controlli
    studenti_raw: list[str] = field(default_factory=list)  # come scritti nel foglio
    note: str = ""
    riga: int = 0


@dataclass
class Abbinamento:
    """Lezione già decisa: quel docente con quel ragazzo (o con un laboratorio LMI) in quel giorno e a quell'ora."""
    docente: str
    studente: str = ""        # id "COGNOME NOME", risolto dai controlli
    studente_raw: str = ""    # come scritto nel foglio: un ragazzo oppure il nome di un laboratorio LMI
    tipo: str = ""            # TIPO_STRUM1 / TIPO_STRUM2 / TIPO_LMC / TIPO_LMI; vuoto = da dedurre
    fascia: int = -1          # 0..19
    note: str = ""
    riga: int = 0
    laboratorio: str = ""     # nome del laboratorio, se è un LMI
    gruppo: int | None = None  # numero del gruppo di musica da camera, se è un LMC di gruppo
    studenti: list[str] = field(default_factory=list)  # id dei ragazzi del gruppo o del laboratorio


@dataclass
class LaboratorioLMI:
    nome: str
    classi: str
    docente: str
    aula: str
    giorno_ora: str
    studenti: str            # "Cognome Nome, Cognome Nome, …" come scritto nel foglio
    note: str = ""
    riga: int = 0
    studenti_id: list[str] = field(default_factory=list)  # risolti dai controlli, se serve collocarlo

    @property
    def nomi_studenti(self) -> list[str]:
        return [n.strip() for n in (self.studenti or "").replace(";", ",").split(",") if n.strip()]


@dataclass
class Parametri:
    max_rientri: int = 2
    max_rientri_vicini: int = 3
    soglia_km_vicino: float = 5.0
    soglia_min_vicino: int = 25  # minuti di ritorno a casa sotto i quali si è "vicini" (se ci sono i trasporti)
    timeout_s: int = 120
    indirizzo_scuola: str = ""


@dataclass
class DatiInput:
    studenti: list[Studente]
    docenti: list[Docente]
    gruppi: list[GruppoLMC]
    lmi: list[LaboratorioLMI]
    parametri: Parametri
    abbinamenti: list[Abbinamento] = field(default_factory=list)
    percorso: Path | None = None
    avvisi: list["Problema"] = field(default_factory=list)  # segnalazioni non bloccanti della lettura

    def docente(self, nome: str) -> Docente | None:
        for d in self.docenti:
            if d.nome == nome:
                return d
        return None

    def studente(self, id_: str) -> Studente | None:
        for s in self.studenti:
            if s.id == id_:
                return s
        return None

    def gruppo_numero(self, numero: int) -> GruppoLMC | None:
        for g in self.gruppi:
            if g.numero == numero:
                return g
        return None

    def laboratorio(self, nome: str) -> LaboratorioLMI | None:
        cercato = " ".join((nome or "").upper().split())
        for l in self.lmi:
            if " ".join(l.nome.upper().split()) == cercato:
                return l
        return None

    def gruppo_di(self, id_studente: str) -> GruppoLMC | None:
        for g in self.gruppi:
            if id_studente in g.studenti:
                return g
        return None


# ── Problemi (controlli e motore) ────────────────────────────────────────────

@dataclass
class Problema:
    livello: str  # "errore" oppure "avviso"
    dove: str  # es. "Studenti, riga 12" oppure "Docenti: Simonelli"
    messaggio: str
    categoria: str = ""  # per raggruppare gli avvisi nel riepilogo prima del calcolo

    def __str__(self) -> str:
        return f"[{self.livello.upper()}] {self.dove}: {self.messaggio}"


class ProblemiError(Exception):
    """Il calcolo si è fermato: l'elenco dei problemi va mostrato all'utente."""

    def __init__(self, problemi: list[Problema]):
        self.problemi = problemi
        super().__init__("\n".join(str(p) for p in problemi))

    @property
    def errori(self) -> list[Problema]:
        return [p for p in self.problemi if p.livello == "errore"]

    @property
    def avvisi(self) -> list[Problema]:
        return [p for p in self.problemi if p.livello == "avviso"]


# ── Output ───────────────────────────────────────────────────────────────────

@dataclass
class Lezione:
    docente: str
    fascia: int  # 0..19
    tipo: str  # TIPO_STRUM1 / TIPO_STRUM2 / TIPO_LMC / TIPO_ACCOMP
    studenti: list[str] = field(default_factory=list)  # id studente; 1 per S1/S2, 2-5 per LMC, 0 per ACC
    gruppo: int | None = None  # numero gruppo LMC
    due_ore: bool = False  # seconda ora consecutiva della stessa lezione (S1 con ore_consecutive)
    nome: str = ""  # nome del laboratorio, per le lezioni LMI

    @property
    def giorno(self) -> int:
        return giorno_ora(self.fascia)[0]

    @property
    def ora(self) -> int:
        return giorno_ora(self.fascia)[1]

    @property
    def studente(self) -> str | None:
        return self.studenti[0] if len(self.studenti) == 1 else None


@dataclass
class Orario:
    lezioni: list[Lezione]
    dati: DatiInput
    stato: str = ""  # "OPTIMAL" / "FEASIBLE"
    secondi: float = 0.0
    avvisi: list[Problema] = field(default_factory=list)  # preferenze non soddisfatte, spiegate

    # ── viste comode per export e anteprima ──
    def per_docente_fascia(self) -> dict[tuple[str, int], Lezione]:
        return {(l.docente, l.fascia): l for l in self.lezioni}

    def lezioni_di(self, id_studente: str) -> list[Lezione]:
        return sorted((l for l in self.lezioni if id_studente in l.studenti), key=lambda l: l.fascia)

    def rientri_di(self, id_studente: str) -> int:
        return len({l.giorno for l in self.lezioni_di(id_studente)})

    def buchi_di(self, id_studente: str) -> int:
        tot = 0
        per_giorno: dict[int, list[int]] = {}
        for l in self.lezioni_di(id_studente):
            per_giorno.setdefault(l.giorno, []).append(l.ora)
        for ore in per_giorno.values():
            tot += (max(ore) - min(ore) + 1) - len(ore)
        return tot

    def lezioni_docente(self, docente: str) -> list[Lezione]:
        return sorted((l for l in self.lezioni if l.docente == docente), key=lambda l: l.fascia)

    def e_lmc(self, l: Lezione) -> bool:
        return l.tipo == TIPO_LMC

    def e_accomp(self, l: Lezione) -> bool:
        return l.tipo == TIPO_ACCOMP
