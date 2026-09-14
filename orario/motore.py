"""Motore di calcolo dell'orario: modello CP-SAT (OR-Tools).

Trasforma un `DatiInput` già controllato in un `Orario`. Le regole rigide (sez. 6 del
documento di progetto) sono vincoli del modello; le preferenze entrano nella funzione
obiettivo con pesi decrescenti: buchi, rientri in più, distanza progressiva, compattezza
dei docenti.

Regola aggiuntiva: le ore di accompagnamento di un docente stanno sempre in coda alle sue
lezioni della stessa giornata (una A fissata a mano prevale: le lezioni vanno prima).

Con `precedente` (l'orario della volta prima) il motore ricalcola cambiando il meno possibile:
ogni lezione spostata costa più di qualunque preferenza.

Se qualche lezione non è collocabile il calcolo si ferma con `ProblemiError` (niente orari
parziali): si rilancia un modello "diagnostico" che può lasciare lezioni fuori e si spiega,
per ciascuna, perché non ci sta.
"""

from __future__ import annotations

import os

import time
from collections import Counter, defaultdict
from collections.abc import Callable
from dataclasses import dataclass, field

from ortools.sat.python import cp_model

from .costanti import (
    GIORNI_LUNGHI, N_FASCE, N_GIORNI, N_ORE, ORE, ORE_FINE, TIPO_ACCOMP, TIPO_LMC, TIPO_LMI, TIPO_STRUM1,
    TIPO_STRUM2, fascia, giorno_ora, nome_fascia,
)
from .modello import DatiInput, Docente, Lezione, Orario, Problema, ProblemiError, Studente

# ── Pesi della funzione obiettivo (ordine di importanza decrescente) ──────────
PESO_BUCO = 1000              # per ogni fascia vuota tra due lezioni dello studente
PESO_RIENTRO_BASE = 600       # per ogni rientro oltre max_rientri: × (1 + 2·km_norm)
PESO_RIENTRO_VICINO = 150     # idem, per chi abita vicino
PESO_ORA_TARDIVA = 12         # × ora, per tutti: si riempiono prima le prime fasce del pomeriggio
PESO_DIST_LONTANO = 200       # × lontananza × ora: chi abita lontano paga caro le ore tarde
PESO_VICINO_PRESTO = 150      # × (1 − lontananza) × ore che mancano alla fine: chi abita vicino paga le
                              # ore presto, così le prime ore restano a chi viene da fuori. Insieme
                              # valgono meno di un buco: l'ordine per distanza non crea vuoti a nessuno.
PESO_MINUTO_ARRIVO = 1        # × minuti di ritardo dell'arrivo a casa rispetto al meglio che quel ragazzo può
                              # fare, × (1 + lontananza). Conta l'ORA IN CUI ARRIVA, non la durata del viaggio:
                              # se il treno è sempre quello, finire più tardi non lo fa arrivare più tardi, e
                              # prima si contava solo l'attesa, che premiava l'ultima ora.
PESO_MINUTO_SERA = 2          # × minuti di arrivo a casa dopo le 17:30, per tutti e in valore assoluto:
                              # è il danno vero (cena, compiti, buio). Serve anche a dare le prime ore a
                              # chi ha i mezzi peggiori: il termine relativo, misurato sul meglio che quel
                              # ragazzo può fare, non distingueva chi arriva comunque tardissimo.
PESO_ULTIMA_LONTANO = 400     # × lontananza: chi abita lontano evita l'ultima fascia, anche quando il viaggio
                              # sarebbe identico (uscire alle 17:30 da lontano è comunque peggio)
PESO_GIORNI_VICINI = 4000     # le 2 ore di 1° strumento nello stesso giorno o in due giorni di fila.
                              # Sopra ogni altro peso, e sopra il costo di spostarle entrambe in un
                              # ricalcolo: si accetta un rientro in più pur di distanziarle. Resta però
                              # cedevole, perché con la regola rigida certi orari non esisterebbero.
PESO_RIENTRO_LONTANO = 350    # × lontananza, per ogni pomeriggio oltre il primo: chi viene da lontano
                              # concentra tutto in un giorno anche restando sotto il massimo dei rientri
PESO_SENZA_MEZZO = 3000       # lezione in una fascia dopo la quale non c'è un mezzo per tornare a casa
PESO_ATTESA_INIZIO = 120      # × ore di attesa fra la fine del mattino e la prima lezione del pomeriggio,
                              # × (0,3 + 1,7 × lontananza): aspettare a vuoto pesa a chi non può tornare
                              # a casa nel frattempo; chi abita a dieci minuti ci torna
PESO_BUCO_DOCENTE = 300       # per ogni fascia vuota tra due impegni del docente (sotto i buchi studente)
PESO_LMC_ATTACCATA = 700      # ora individuale non nello stesso pomeriggio della musica da camera
                              # (sotto il peso di un buco: non conviene creare un vuoto pur di attaccarla)
PESO_SPOSTAMENTO = 800        # ricalcolo: lezione spostata rispetto all'orario precedente. Sotto il peso
                              # di un buco: adattare l'orario esistente non deve mai crearne uno

NUM_WORKERS = max(4, os.cpu_count() or 4)  # usa tutti i core della macchina (anche su Windows)
TIMEOUT_DIAGNOSI_S = 30

_NOME_TIPO = {
    TIPO_STRUM1: "1° strumento",
    TIPO_STRUM2: "2° strumento",
    TIPO_LMC: "musica da camera",
    TIPO_LMI: "laboratorio d'insieme",
    TIPO_ACCOMP: "accompagnamento",
}


# ── Unità da collocare ───────────────────────────────────────────────────────

@dataclass
class _Unita:
    idx: int
    docente: str
    tipo: str
    studenti: list[str] = field(default_factory=list)
    gruppo: int | None = None
    fasce: list[int] = field(default_factory=list)   # fasce ammesse (dominio)
    seguente: int | None = None   # idx della seconda ora consecutiva (solo sulla prima)
    seconda: bool = False         # è la seconda ora di una coppia consecutiva
    nome: str = ""                # nome del laboratorio, per le unità LMI
    fissata: bool = False         # bloccata in una fascia dal foglio degli abbinamenti fissi

    def descrizione(self, dati: DatiInput, n: int = 1) -> str:
        if self.tipo == TIPO_ACCOMP:
            return (f"ora di accompagnamento del docente {self.docente}" if n == 1
                    else f"{n} ore di accompagnamento del docente {self.docente}")
        if self.tipo == TIPO_LMC:
            return f"lezione di musica da camera del gruppo {self.gruppo} con il docente {self.docente}"
        if self.tipo == TIPO_LMI:
            return f"laboratorio {self.nome} con il docente {self.docente}"
        return (f"ora di {_NOME_TIPO[self.tipo]} con il docente {self.docente}" if n == 1
                else f"{n} ore di {_NOME_TIPO[self.tipo]} con il docente {self.docente}")


def _raggruppa(unita: list[_Unita]) -> list[tuple[_Unita, int]]:
    """Unità identiche (stesso docente, tipo, studenti) → (rappresentante, quante)."""
    conta: dict[tuple, int] = {}
    prima: dict[tuple, _Unita] = {}
    for u in unita:
        k = (u.docente, u.tipo, frozenset(u.studenti))
        conta[k] = conta.get(k, 0) + 1
        prima.setdefault(k, u)
    return [(prima[k], conta[k]) for k in prima]


def _blocca_abbinamenti(dati: DatiInput, unita: list[_Unita]) -> None:
    """Fissa nella loro fascia le lezioni già decise nel foglio degli abbinamenti.

    Restringe il dominio dell'unità a quella sola fascia; se un ragazzo ha due ore dello stesso
    tipo con lo stesso docente, si blocca la prima ancora libera.
    """
    bloccate: set[int] = set()
    problemi: list[Problema] = []
    for ab in dati.abbinamenti:
        if ab.tipo == TIPO_LMI and ab.fascia >= 0 and ab.studenti:
            # il laboratorio non fa parte delle ore da distribuire: esiste solo perché è stato fissato qui
            unita.append(_Unita(len(unita), ab.docente, TIPO_LMI, list(ab.studenti),
                                fasce=[ab.fascia], nome=ab.laboratorio, fissata=True))
            continue
        if not ab.studente or not ab.tipo or ab.fascia < 0:
            continue
        candidate = [u for u in unita
                     if u.idx not in bloccate and u.tipo == ab.tipo and u.docente == ab.docente
                     and ab.studente in u.studenti and not u.seconda]
        if not candidate:
            problemi.append(Problema(
                "errore", f"Abbinamenti fissi, riga {ab.riga}",
                f"non resta nessuna lezione da fissare per {ab.studente.title()} con {ab.docente}: "
                "forse ci sono più abbinamenti dello stesso tipo di quante sono le sue ore."))
            continue
        u = candidate[0]
        fasce = [ab.fascia] if ab.fascia in u.fasce else []
        if u.seguente is not None:  # 2 ore consecutive: serve anche la fascia dopo
            seconda = next(v for v in unita if v.idx == u.seguente)
            if ab.fascia + 1 not in seconda.fasce or giorno_ora(ab.fascia)[1] == N_ORE - 1:
                fasce = []
        if not fasce:
            problemi.append(Problema(
                "errore", f"Abbinamenti fissi, riga {ab.riga}",
                f"{ab.studente.title()} con {ab.docente} {nome_fascia(ab.fascia)} non è possibile: "
                "il docente non è libero in quell'ora, oppure servono due ore consecutive che non ci stanno."))
            continue
        u.fasce = fasce
        u.fissata = True
        bloccate.add(u.idx)
        if u.seguente is not None:
            seconda = next(v for v in unita if v.idx == u.seguente)
            seconda.fasce = [ab.fascia + 1]
            seconda.fissata = True
            bloccate.add(seconda.idx)
    if problemi:
        raise ProblemiError(problemi)


def _crea_unita(dati: DatiInput) -> list[_Unita]:
    docenti = {d.nome: d for d in dati.docenti}
    studenti = {s.id: s for s in dati.studenti}
    unita: list[_Unita] = []

    def dominio(nome_doc: str, ids: list[str], accomp: bool) -> list[int]:
        d = docenti[nome_doc]
        vietate: set[int] = set()
        for i in ids:
            vietate |= studenti[i].fasce_vietate     # giorni interi più impegni personali
        out = []
        for f, v in d.disponibilita.items():
            if (v == "X" or (accomp and v == "A")) and f not in vietate:
                out.append(f)
        return sorted(out)

    def nuova(nome_doc: str, tipo: str, ids: list[str], gruppo: int | None = None) -> _Unita:
        u = _Unita(len(unita), nome_doc, tipo, list(ids), gruppo,
                   fasce=dominio(nome_doc, ids, tipo == TIPO_ACCOMP))
        unita.append(u)
        return u

    for s in dati.studenti:
        h1, h2, _ = s.ore
        prime = [nuova(s.doc1, TIPO_STRUM1, [s.id]) for _ in range(h1)]
        if s.ore_consecutive and h1 == 2:
            a, b = prime
            a.seguente, b.seconda = b.idx, True
            dom = set(a.fasce)
            a.fasce = [f for f in dom if giorno_ora(f)[1] < N_ORE - 1 and (f + 1) in dom]
            b.fasce = [f + 1 for f in a.fasce]
        for _ in range(h2):
            nuova(s.doc2, TIPO_STRUM2, [s.id])
    for g in dati.gruppi:
        nuova(g.docente, TIPO_LMC, g.studenti, g.numero)
    for d in dati.docenti:
        for _ in range(d.ore_accomp):
            nuova(d.nome, TIPO_ACCOMP, [])
    return unita


# ── Km normalizzati ──────────────────────────────────────────────────────────

_MINUTI_FINE = [int(o[:2]) * 60 + int(o[3:5]) for o in ORE_FINE]
ARRIVO_SERA = 17 * 60 + 30    # oltre quest'ora l'arrivo a casa comincia a pesare in assoluto


def _arrivo_a_casa(t, o: int) -> int | None:
    """Minuti dopo mezzanotte a cui il ragazzo arriva a casa se la lezione finisce nella fascia `o`."""
    testo = t.arrivi[o] if o < len(t.arrivi) else ""
    ore_min = str(testo).strip().split(":")
    if len(ore_min) >= 2 and ore_min[0].isdigit() and ore_min[1][:2].isdigit():
        return int(ore_min[0]) * 60 + int(ore_min[1][:2])
    minuti = t.minuti[o] if o < len(t.minuti) else None
    return None if minuti is None else _MINUTI_FINE[o] + 5 + minuti


def _lontananza(s: Studente) -> float:
    """Quanto è 'lontano' uno studente, in minuti di ritorno a casa.

    Con i dati dei mezzi: il ritorno migliore tra le 4 fasce. Senza: stima dai km (10 min + 2,5 min/km).
    Nulla di nulla → 0 (vicino)."""
    return s.minuti_ritorno


def _km_norm(dati: DatiInput) -> dict[str, float]:
    """Rango normalizzato della lontananza in [0, 1] (0 = il più vicino); pari merito → rango medio."""
    valori = sorted((_lontananza(s), s.id) for s in dati.studenti)
    n = len(valori)
    out: dict[str, float] = {}
    i = 0
    while i < n:
        j = i
        while j + 1 < n and valori[j + 1][0] == valori[i][0]:
            j += 1
        rango = (i + j) / 2
        for k in range(i, j + 1):
            out[valori[k][1]] = rango / (n - 1) if n > 1 else 0.0
        i = j + 1
    return out


def _vicino(s: Studente, dati: DatiInput) -> bool:
    if s.trasporto is not None and s.trasporto.minuti_min is not None:
        return s.trasporto.minuti_min < dati.parametri.soglia_min_vicino
    return s.km is None or s.km < dati.parametri.soglia_km_vicino


def _etichetta(s: Studente, con_km: bool = False) -> str:
    testo = f"{s.cognome.title()} {s.nome.title()} ({s.classe}ª"
    if con_km and s.trasporto is not None and s.trasporto.minuti_min is not None:
        testo += f", {s.trasporto.minuti_min} min di ritorno"
    elif con_km and s.km is not None:
        testo += f", {s.km:g} km"
    return testo + ")"


def _giorno(g: int) -> str:
    return GIORNI_LUNGHI[g].lower()


# ── Modello CP-SAT ───────────────────────────────────────────────────────────

class _Modello:
    """Costruisce il modello. Con `diagnosi=True` ogni unità può restare fuori (variabile `manca`)."""

    def __init__(self, dati: DatiInput, unita: list[_Unita], diagnosi: bool = False,
                 precedente: dict[int, int] | None = None):
        self.dati = dati
        self.unita = unita
        self.diagnosi = diagnosi
        self.precedente = precedente or {}   # idx unità → fascia dell'orario precedente
        self.m = cp_model.CpModel()
        self.x: dict[tuple[int, int], cp_model.IntVar] = {}
        self.manca: dict[int, cp_model.IntVar] = {}
        self.occ: dict[tuple[str, int], cp_model.IntVar] = {}    # studente, fascia
        self.docc: dict[tuple[str, int], cp_model.IntVar] = {}   # docente, fascia
        self.buchi_stud: dict[str, list] = defaultdict(list)
        self.buchi_doc: dict[str, list] = defaultdict(list)
        self.rientri: dict[str, object] = {}
        self.costi: list[tuple[int, object]] = []
        self._costruisci()

    # -- costruzione --
    def _costruisci(self) -> None:
        m, dati, unita = self.m, self.dati, self.unita
        par = dati.parametri
        km_norm = _km_norm(dati)

        # variabili x[u, f]
        for u in unita:
            for f in u.fasce:
                self.x[u.idx, f] = m.new_bool_var(f"x_{u.idx}_{f}")
            somma = sum(self.x[u.idx, f] for f in u.fasce)
            if self.diagnosi:
                self.manca[u.idx] = m.new_bool_var(f"manca_{u.idx}")
                m.add(somma + self.manca[u.idx] == 1)
            else:
                m.add(somma == 1)

        # ore consecutive: seconda ora = prima ora + 1
        for u in unita:
            if u.seguente is not None:
                b = unita[u.seguente]
                for f in u.fasce:
                    m.add(self.x[u.idx, f] == self.x[b.idx, f + 1])

        # docenti: al massimo 1 unità per fascia; A → esattamente un'ora di accompagnamento
        per_doc: dict[str, list[_Unita]] = defaultdict(list)
        for u in unita:
            per_doc[u.docente].append(u)
        for d in dati.docenti:
            unita_d = per_doc.get(d.nome, [])
            for f in range(N_FASCE):
                presenti = [self.x[u.idx, f] for u in unita_d if (u.idx, f) in self.x]
                if not presenti:
                    continue
                v = m.new_bool_var(f"doc_{d.riga}_{f}")
                m.add(sum(presenti) == v)
                self.docc[d.nome, f] = v
                if d.disponibilita.get(f) == "A":
                    acc = [self.x[u.idx, f] for u in unita_d if u.tipo == TIPO_ACCOMP and (u.idx, f) in self.x]
                    m.add(sum(acc) == 1)
            # accompagnamento sempre DOPO le lezioni della stessa giornata
            acc_u = [u for u in unita_d if u.tipo == TIPO_ACCOMP]
            lez_u = [u for u in unita_d if u.tipo != TIPO_ACCOMP]
            if acc_u and lez_u:
                for g in range(N_GIORNI):
                    for o1 in range(N_ORE - 1):
                        f1 = fascia(g, o1)
                        acc_f1 = [self.x[u.idx, f1] for u in acc_u if (u.idx, f1) in self.x]
                        if not acc_f1:
                            continue
                        for o2 in range(o1 + 1, N_ORE):
                            f2 = fascia(g, o2)
                            lez_f2 = [self.x[u.idx, f2] for u in lez_u if (u.idx, f2) in self.x]
                            if lez_f2:
                                m.add(sum(acc_f1) + sum(lez_f2) <= 1)
            # simmetria: le ore di accompagnamento sono identiche → in ordine crescente di fascia
            if not self.diagnosi:
                for a, b in zip(acc_u, acc_u[1:]):
                    m.add(self._pos(a) < self._pos(b))

        # studenti: al massimo 1 unità per fascia, giorni, rientri, buchi, distanza
        per_stud: dict[str, list[_Unita]] = defaultdict(list)
        for u in unita:
            for sid in u.studenti:
                per_stud[sid].append(u)
        for s in dati.studenti:
            unita_s = per_stud.get(s.id, [])
            occ_s: list = []
            for f in range(N_FASCE):
                presenti = [self.x[u.idx, f] for u in unita_s if (u.idx, f) in self.x]
                if presenti:
                    v = m.new_bool_var(f"occ_{s.riga}_{f}")
                    m.add(sum(presenti) == v)
                    self.occ[s.id, f] = v
                    occ_s.append(v)
                else:
                    occ_s.append(0)
            # simmetria: due ore di 1° strumento non consecutive sono intercambiabili.
            # Non si applica se una delle due è stata fissata a mano: l'ordine lo decide il foglio.
            s1 = [u for u in unita_s if u.tipo == TIPO_STRUM1 and u.seguente is None and not u.seconda]
            if not self.diagnosi and not any(u.fissata for u in s1):
                for a, b in zip(s1, s1[1:]):
                    m.add(self._pos(a) < self._pos(b))
            # le 2 ore di 1° strumento in due giorni NON consecutivi (almeno un giorno in mezzo),
            # salvo un SI esplicito o un abbinamento fisso: se l'ora l'ha scelta una persona, vince lei.
            # Vietando ogni coppia di giorni vicini (g, g+1) si escludono sia lo stesso giorno sia il
            # giorno dopo: restano lunedì-mercoledì, lunedì-giovedì, lunedì-venerdì, martedì-giovedì…
            # vale anche per chi abita lontano: due lezioni dello stesso strumento troppo vicine non
            # hanno senso, e su questo la distanza da casa non conta
            if (s.primo_separato and len(s1) == 2 and not s.giorno_unico
                    and not any(u.fissata for u in s1)):
                for g in range(N_GIORNI - 1):
                    vicini = [self.x[u.idx, f] for u in s1
                              for f in range(fascia(g, 0), fascia(g + 1, 0) + N_ORE)
                              if (u.idx, f) in self.x]
                    if len(vicini) > 1:
                        viola = m.new_bool_var(f"vicine_{s.riga}_{g}")
                        m.add(sum(vicini) - 1 <= viola)
                        self.costi.append((PESO_GIORNI_VICINI, viola))
            # le ore individuali stanno volentieri nello stesso pomeriggio della musica da camera:
            # senza buchi, "stesso giorno" vuol dire attaccate, prima o dopo la lezione di gruppo
            u_lmc = next((u for u in unita_s if u.tipo == TIPO_LMC), None)
            individuali = [u for u in unita_s if u.tipo in (TIPO_STRUM1, TIPO_STRUM2)]
            if u_lmc is not None and individuali:
                for u in individuali:
                    insieme = []
                    for g in range(N_GIORNI):
                        fasce_g = [fascia(g, o) for o in range(N_ORE)]
                        a = [self.x[u_lmc.idx, f] for f in fasce_g if (u_lmc.idx, f) in self.x]
                        b = [self.x[u.idx, f] for f in fasce_g if (u.idx, f) in self.x]
                        if not a or not b:
                            continue
                        v = m.new_bool_var(f"con_lmc_{s.riga}_{u.idx}_{g}")
                        m.add(v <= sum(a))
                        m.add(v <= sum(b))
                        m.add(v >= sum(a) + sum(b) - 1)
                        insieme.append(v)
                    if insieme:
                        attaccata = m.new_bool_var(f"attaccata_{s.riga}_{u.idx}")
                        m.add(attaccata == sum(insieme))
                        self.costi.append((PESO_LMC_ATTACCATA, 1 - attaccata))

            # giorni con almeno una lezione
            giorni_var: list = []
            for g in range(N_GIORNI):
                occ_g = [occ_s[fascia(g, o)] for o in range(N_ORE)]
                variabili = [v for v in occ_g if not isinstance(v, int)]
                if not variabili:
                    continue
                dv = m.new_bool_var(f"giorno_{s.riga}_{g}")
                for v in variabili:
                    m.add(dv >= v)
                m.add(dv <= sum(variabili))
                giorni_var.append(dv)
                # attesa fra la fine del mattino e la prima lezione: il pomeriggio cominci presto
                peso_attesa = round(PESO_ATTESA_INIZIO * (0.3 + 1.7 * km_norm[s.id]))
                for o in range(N_ORE - 1):
                    prima = [v for v in occ_g[:o + 1] if not isinstance(v, int)]
                    aspetta = m.new_bool_var(f"attesa_{s.riga}_{g}_{o}")
                    m.add(aspetta >= dv - sum(prima)) if prima else m.add(aspetta >= dv)
                    self.costi.append((peso_attesa, aspetta))
                # buchi nel pomeriggio
                self.buchi_stud[s.id].extend(self._buchi(occ_g, f"bs_{s.riga}_{g}"))
            rientri = sum(giorni_var) if giorni_var else 0
            self.rientri[s.id] = rientri
            if giorni_var and km_norm[s.id]:
                # ogni pomeriggio in più costa a chi abita lontano, anche sotto il massimo dei rientri
                oltre_il_primo = m.new_int_var(0, len(giorni_var), f"gg_{s.riga}")
                m.add(oltre_il_primo >= rientri - 1)
                self.costi.append((round(PESO_RIENTRO_LONTANO * km_norm[s.id]), oltre_il_primo))
            if giorni_var:
                m.add(rientri <= par.max_rientri_vicini)
                if s.giorno_unico:
                    m.add(rientri <= 1)
                tetto = max(0, par.max_rientri_vicini - par.max_rientri)
                if tetto > 0:
                    extra = m.new_int_var(0, tetto, f"extra_{s.riga}")
                    m.add(extra >= rientri - par.max_rientri)
                    if _vicino(s, dati):
                        peso = PESO_RIENTRO_VICINO
                    else:
                        peso = round(PESO_RIENTRO_BASE * (1 + 2 * km_norm[s.id]))
                    self.costi.append((peso, extra))
            # distanza progressiva: prime fasce piene, i lontani ancora più presto
            kn = km_norm[s.id]
            tr = s.trasporto
            arrivi = [_arrivo_a_casa(tr, o) for o in range(N_ORE)] if tr is not None and tr.valido else []
            validi = [a for a in arrivi if a is not None]
            prima_possibile = min(validi) if validi else None
            for f in range(N_FASCE):
                v = occ_s[f]
                if isinstance(v, int):
                    continue
                o = giorno_ora(f)[1]
                # criterio principale (uguale con km o con mezzi): le prime ore a chi abita lontano,
                # le ultime a chi abita vicino
                c = round((PESO_ORA_TARDIVA + PESO_DIST_LONTANO * kn) * o
                          + PESO_VICINO_PRESTO * (1 - kn) * (N_ORE - 1 - o))
                if o == N_ORE - 1 and kn:
                    c += round(PESO_ULTIMA_LONTANO * kn)   # i lontani saltano l'ultima ora
                if tr is not None and tr.valido:
                    # con i dati dei mezzi: conta a che ora arriva a casa, non quanto dura il viaggio.
                    # Il riferimento è il meglio che quel ragazzo può fare, così la fascia migliore costa 0.
                    m_o = tr.minuti[o] if o < len(tr.minuti) else None
                    if m_o is None:
                        c += PESO_SENZA_MEZZO     # nessun mezzo per tornare a casa dopo quella fascia
                    else:
                        casa = _arrivo_a_casa(tr, o)
                        if casa is not None and prima_possibile is not None:
                            c += round(PESO_MINUTO_ARRIVO * (1 + kn) * max(0, casa - prima_possibile))
                            c += PESO_MINUTO_SERA * max(0, casa - ARRIVO_SERA)
                if c:
                    self.costi.append((c, v))
            for b in self.buchi_stud[s.id]:
                self.costi.append((PESO_BUCO, b))

        # compattezza docenti
        for d in dati.docenti:
            for g in range(N_GIORNI):
                occ_g = [self.docc.get((d.nome, fascia(g, o)), 0) for o in range(N_ORE)]
                self.buchi_doc[d.nome].extend(self._buchi(occ_g, f"bd_{d.riga}_{g}"))
            for b in self.buchi_doc[d.nome]:
                self.costi.append((PESO_BUCO_DOCENTE, b))

        # ricalcolo: restare nella fascia dell'orario precedente
        self.spostata: dict[int, object] = {}
        if not self.diagnosi:
            for idx, f in self.precedente.items():
                if (idx, f) in self.x:
                    m.add_hint(self.x[idx, f], 1)
                    self.costi.append((PESO_SPOSTAMENTO, 1 - self.x[idx, f]))

        if self.diagnosi:
            m.minimize(sum(self.manca.values()))
        else:
            m.minimize(sum(c * v for c, v in self.costi))

    def _pos(self, u: _Unita):
        return sum(f * self.x[u.idx, f] for f in u.fasce)

    def _buchi(self, occ_g: list, nome: str) -> list:
        """Variabili 'buco' per un pomeriggio: fascia vuota con lezioni sia prima che dopo.

        occ_g: 4 elementi, variabile booleana oppure 0 se la fascia è impossibile.
        """
        m = self.m
        out = []
        for o in range(1, N_ORE - 1):
            prima = [v for v in occ_g[:o] if not isinstance(v, int)]
            dopo = [v for v in occ_g[o + 1:] if not isinstance(v, int)]
            if not prima or not dopo:
                continue
            lt = self._or(prima, f"{nome}_lt{o}")
            gt = self._or(dopo, f"{nome}_gt{o}")
            b = m.new_bool_var(f"{nome}_b{o}")
            m.add(b >= lt + gt - 1 - occ_g[o])
            out.append(b)
        return out

    def _or(self, variabili: list, nome: str):
        if len(variabili) == 1:
            return variabili[0]
        v = self.m.new_bool_var(nome)
        for w in variabili:
            self.m.add(v >= w)
        return v

    # -- lettura soluzione --
    def fascia_di(self, solver: cp_model.CpSolver, u: _Unita) -> int | None:
        for f in u.fasce:
            if solver.value(self.x[u.idx, f]):
                return f
        return None


class _Avanzamento(cp_model.CpSolverSolutionCallback):
    def __init__(self, progresso: Callable[[str], None], diagnosi: bool):
        super().__init__()
        self.progresso = progresso
        self.diagnosi = diagnosi
        self.n = 0
        self._ultimo = 0.0

    def on_solution_callback(self) -> None:
        self.n += 1
        ora = time.monotonic()
        if ora - self._ultimo < 2.0:
            return
        self._ultimo = ora
        if self.diagnosi:
            self.progresso(f"Diagnosi: lezioni fuori posto {int(self.objective_value)}…")
        else:
            self.progresso(f"Soluzione n. {self.n} trovata (costo {int(self.objective_value)}), continuo a migliorare…")


def _risolvi(modello: _Modello, timeout: float, progresso: Callable[[str], None]):
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = max(1.0, float(timeout))
    solver.parameters.num_workers = NUM_WORKERS
    stato = solver.solve(modello.m, _Avanzamento(progresso, modello.diagnosi))
    return solver, stato


# ── Funzione principale ──────────────────────────────────────────────────────

def calcola(dati: DatiInput, progresso: Callable[[str], None] | None = None,
            precedente: list[Lezione] | None = None) -> Orario:
    """Calcola l'orario. Solleva `ProblemiError` se qualche lezione non è collocabile.

    `precedente`: lezioni dell'orario calcolato la volta prima. Se fornite, ogni lezione resta
    nella sua fascia salvo quando le regole rigide lo impediscono (ricalcolo a modifiche minime).
    """
    progresso = progresso or (lambda _msg: None)
    inizio = time.monotonic()
    timeout = dati.parametri.timeout_s

    # i cognomi scritti nei gruppi di musica da camera vengono tradotti in studenti da controlli.controlla:
    # senza quel passaggio i gruppi risulterebbero vuoti e l'orario sarebbe sbagliato in silenzio
    non_risolti = [g for g in dati.gruppi if g.studenti_raw and not g.studenti]
    if non_risolti:
        raise ProblemiError([Problema(
            "errore", "Programma",
            f"I gruppi di musica da camera non sono stati controllati ({len(non_risolti)} gruppi senza studenti "
            "riconosciuti): va eseguito 'controlli.controlla(dati)' prima del calcolo.")])

    progresso("Costruzione del modello…")
    unita = _crea_unita(dati)
    _blocca_abbinamenti(dati, unita)

    # unità senza nessuna fascia possibile: inutile far partire il calcolo
    vuote = [u for u in unita if not u.fasce]
    if vuote:
        problemi = [_spiega_unita(dati, unita, u, None, n) for u, n in _raggruppa(vuote)]
        raise ProblemiError(problemi)

    fasce_prec = _abbina_precedente(unita, precedente) if precedente else {}
    modello = _Modello(dati, unita, precedente=fasce_prec)
    progresso(f"Ricerca soluzione (fino a {timeout} s)…")
    solver, stato = _risolvi(modello, timeout, progresso)

    if stato not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        progresso("Nessuna combinazione trovata: cerco il motivo…")
        raise ProblemiError(_diagnosi(dati, unita, stato, progresso))

    progresso("Soluzione trovata, verifica…")
    lezioni = _estrai_lezioni(modello, solver, unita)
    _verifica(dati, unita, lezioni)
    orario = Orario(
        lezioni=lezioni, dati=dati,
        stato="OPTIMAL" if stato == cp_model.OPTIMAL else "FEASIBLE",
        secondi=round(time.monotonic() - inizio, 1),
    )
    orario.avvisi = _avvisi(orario)
    if precedente:
        avvisi_sp, riassunto = _avvisi_spostamenti(dati, unita, lezioni, fasce_prec)
        orario.avvisi = avvisi_sp + orario.avvisi
        progresso(f"Fatto. {riassunto}")
    else:
        progresso("Fatto.")
    return orario


def _chiave(docente: str, tipo: str, studenti: list[str]) -> tuple:
    return (docente, tipo, frozenset(studenti))


def _abbina_precedente(unita: list[_Unita], precedente: list[Lezione]) -> dict[int, int]:
    """idx unità → fascia dell'orario precedente, per chiave (docente, tipo, studenti).

    Unità con la stessa chiave (due ore di 1° strumento, ore di accompagnamento) prendono le
    fasce precedenti in ordine crescente. Lezioni senza unità corrispondente vengono ignorate.
    """
    prec: dict[tuple, list[int]] = defaultdict(list)
    for l in precedente:
        prec[_chiave(l.docente, l.tipo, l.studenti)].append(l.fascia)
    for v in prec.values():
        v.sort()
    gruppi: dict[tuple, list[_Unita]] = defaultdict(list)
    for u in unita:
        gruppi[_chiave(u.docente, u.tipo, u.studenti)].append(u)
    out: dict[int, int] = {}
    for k, lista in gruppi.items():
        for u, f in zip(lista, prec.get(k, [])):
            out[u.idx] = f
    return out


def _avvisi_spostamenti(dati: DatiInput, unita: list[_Unita], lezioni: list[Lezione],
                        fasce_prec: dict[int, int]) -> tuple[list[Problema], str]:
    """Un avviso per lezione spostata rispetto all'orario precedente + riassunto in testa."""
    docenti = {d.nome: d for d in dati.docenti}
    confermate = spostate = 0
    nuove = len(unita) - len(fasce_prec)
    avvisi: list[Problema] = []
    for u, l in zip(unita, lezioni):
        if u.idx not in fasce_prec:
            continue
        fp = fasce_prec[u.idx]
        if l.fascia == fp:
            confermate += 1
            continue
        spostate += 1
        # motivo, quando si deduce dai dati
        d = docenti[u.docente]
        studenti = [s for s in (dati.studente(i) for i in u.studenti) if s]
        g_prec = giorno_ora(fp)[0]
        vp = d.disponibilita.get(fp)
        if vp is None or (u.tipo != TIPO_ACCOMP and vp != "X"):
            motivo = f"il docente non è più disponibile il {_giorno(g_prec)} alle {ORE[giorno_ora(fp)[1]]}"
        elif any(g_prec in s.giorni_non_disp for s in studenti):
            chi = next(s for s in studenti if g_prec in s.giorni_non_disp)
            motivo = (f"{_etichetta(chi)} non è più disponibile il {_giorno(g_prec)}" if len(studenti) > 1
                      else f"lo studente non è più disponibile il {_giorno(g_prec)}")
        elif fp not in u.fasce:
            motivo = "la fascia precedente non è più compatibile con i vincoli (ore consecutive)"
        else:
            motivo = "spostata per far posto alle modifiche"
        if u.tipo == TIPO_ACCOMP:
            chi, cosa, dove = f"Docente {u.docente}", "ora di accompagnamento spostata", f"Docenti: {u.docente}"
        elif u.tipo == TIPO_LMC:
            chi, cosa, dove = f"Gruppo LMC {u.gruppo} (docente {u.docente})", "musica da camera spostata", f"Gruppi LMC, gruppo {u.gruppo}"
        else:
            chi, cosa, dove = _etichetta(studenti[0]), f"{_NOME_TIPO[u.tipo]} spostato", f"Studenti, riga {studenti[0].riga}"
            if u.seconda:
                cosa = f"2ª ora di {_NOME_TIPO[u.tipo]} spostata"
        avvisi.append(Problema("avviso", dove,
                               f"{chi}: {cosa} da {_fascia_breve(fp)} a {_fascia_breve(l.fascia)} ({motivo})"))
    riassunto = (f"Rispetto all'orario precedente: {confermate} lezioni confermate, "
                 f"{spostate} spostate, {nuove} nuove")
    return [Problema("avviso", "Ricalcolo", riassunto)] + avvisi, riassunto


def _fascia_breve(f: int) -> str:
    g, o = giorno_ora(f)
    return f"{GIORNI_LUNGHI[g]} {ORE[o]}"


def _estrai_lezioni(modello: _Modello, solver: cp_model.CpSolver, unita: list[_Unita]) -> list[Lezione]:
    lezioni = []
    for u in unita:
        f = modello.fascia_di(solver, u)
        if f is None:
            raise RuntimeError(f"Errore interno del motore: {u.descrizione(modello.dati)} senza fascia.")
        lezioni.append(Lezione(docente=u.docente, fascia=f, tipo=u.tipo, studenti=list(u.studenti),
                               gruppo=u.gruppo, due_ore=u.seconda, nome=u.nome))
    return lezioni  # stesso ordine delle unità (serve per il confronto con l'orario precedente)


# ── Verifica delle regole rigide sull'orario prodotto ────────────────────────

def _verifica(dati: DatiInput, unita: list[_Unita], lezioni: list[Lezione]) -> None:
    def errore(msg: str) -> None:
        raise RuntimeError(f"Errore interno del motore, orario non valido: {msg}")

    docenti = {d.nome: d for d in dati.docenti}
    studenti = {s.id: s for s in dati.studenti}
    par = dati.parametri
    fissati_s1 = {sid for u in unita if u.fissata and u.tipo == TIPO_STRUM1 for sid in u.studenti}

    if len(lezioni) != len(unita):
        errore(f"{len(lezioni)} lezioni prodotte, {len(unita)} attese.")

    # 1. disponibilità, 2. docente unico per fascia
    per_doc_fascia = Counter((l.docente, l.fascia) for l in lezioni)
    for (d, f), n in per_doc_fascia.items():
        if n > 1:
            errore(f"il docente {d} ha {n} lezioni {nome_fascia(f)}.")
    for l in lezioni:
        v = docenti[l.docente].disponibilita.get(l.fascia)
        if l.tipo == TIPO_ACCOMP:
            if v not in ("X", "A"):
                errore(f"accompagnamento di {l.docente} {nome_fascia(l.fascia)} senza disponibilità.")
        elif v != "X":
            errore(f"lezione di {l.docente} {nome_fascia(l.fascia)} in fascia non disponibile ({v or 'vuota'}).")
        # 4. giorni vietati e impegni personali
        for sid in l.studenti:
            st = studenti[sid]
            if l.giorno in st.giorni_non_disp:
                errore(f"{_etichetta(st)} ha lezione di {_giorno(l.giorno)}, giorno non disponibile.")
            elif l.fascia in st.fasce_non_disp:
                errore(f"{_etichetta(st)} ha lezione {nome_fascia(l.fascia)}, ora in cui ha un impegno.")
    # abbinamenti fissi rispettati
    for ab in dati.abbinamenti:
        if not ab.studente or ab.fascia < 0:
            continue
        if not any(l.fascia == ab.fascia and l.docente == ab.docente and ab.studente in l.studenti
                   for l in lezioni):
            errore(f"l'abbinamento fisso {ab.studente.title()} con {ab.docente} "
                   f"{nome_fascia(ab.fascia)} non è stato rispettato.")

    # A rispettate, numero ACCOMP, accompagnamento in coda alle lezioni della giornata
    for d in dati.docenti:
        acc = [l for l in lezioni if l.docente == d.nome and l.tipo == TIPO_ACCOMP]
        if len(acc) != d.ore_accomp:
            errore(f"il docente {d.nome} ha {len(acc)} ore di accompagnamento invece di {d.ore_accomp}.")
        for f in d.fasce_a:
            if sum(1 for l in acc if l.fascia == f) != 1:
                errore(f"la A di {d.nome} {nome_fascia(f)} non è coperta da un'ora di accompagnamento.")
        lez = [l for l in lezioni if l.docente == d.nome and l.tipo != TIPO_ACCOMP]
        for g in range(N_GIORNI):
            ore_acc = [l.ora for l in acc if l.giorno == g]
            ore_lez = [l.ora for l in lez if l.giorno == g]
            if ore_acc and ore_lez and min(ore_acc) < max(ore_lez):
                errore(f"il docente {d.nome} ha un'ora di accompagnamento prima di una lezione il {_giorno(g)}.")
    # 3, 5, 6, 7 e ore per studente
    for s in dati.studenti:
        mie = [l for l in lezioni if s.id in l.studenti]
        et = _etichetta(s)
        conta = Counter(l.fascia for l in mie)
        for f, n in conta.items():
            if n > 1:
                errore(f"{et} ha {n} lezioni {nome_fascia(f)}.")
        h1, h2, hl = s.ore
        tipi = Counter(l.tipo for l in mie)
        if (tipi[TIPO_STRUM1], tipi[TIPO_STRUM2], tipi[TIPO_LMC]) != (h1, h2, hl):
            errore(f"{et} ha ore {tipi[TIPO_STRUM1]}/{tipi[TIPO_STRUM2]}/{tipi[TIPO_LMC]} invece di {h1}/{h2}/{hl}.")
        for l in mie:
            if l.tipo == TIPO_STRUM1 and l.docente != s.doc1:
                errore(f"{et}: 1° strumento con {l.docente} invece di {s.doc1}.")
            if l.tipo == TIPO_STRUM2 and l.docente != s.doc2:
                errore(f"{et}: 2° strumento con {l.docente} invece di {s.doc2}.")
        if s.ore_consecutive and h1 == 2:
            a, b = sorted((l for l in mie if l.tipo == TIPO_STRUM1), key=lambda l: l.fascia)
            if a.giorno != b.giorno or b.ora != a.ora + 1:
                errore(f"{et}: le due ore di 1° strumento non sono consecutive.")
            if a.due_ore or not b.due_ore:
                errore(f"{et}: indicatore due_ore sbagliato.")

        giorni = {l.giorno for l in mie}
        if s.giorno_unico and len(giorni) > 1:
            errore(f"{et}: giorno unico non rispettato ({len(giorni)} giorni).")
        if len(giorni) > par.max_rientri_vicini:
            errore(f"{et}: {len(giorni)} rientri, oltre il massimo di {par.max_rientri_vicini}.")
    for g in dati.gruppi:
        lg = [l for l in lezioni if l.tipo == TIPO_LMC and l.gruppo == g.numero]
        if len(lg) != 1 or lg[0].docente != g.docente or set(lg[0].studenti) != set(g.studenti):
            errore(f"gruppo LMC {g.numero} non collocato correttamente.")


# ── Avvisi sulle preferenze non soddisfatte ──────────────────────────────────

def _avvisi(orario: Orario) -> list[Problema]:
    dati = orario.dati
    par = dati.parametri
    km_norm = _km_norm(dati)
    avvisi: list[Problema] = []
    for s in dati.studenti:
        if s.ore[0] != 2 or not s.primo_separato or s.giorno_unico:
            continue
        gg = sorted(l.giorno for l in orario.lezioni_di(s.id) if l.tipo == TIPO_STRUM1)
        if len(gg) == 2 and gg[1] - gg[0] < 2:
            come = "nello stesso giorno" if gg[0] == gg[1] else "in due giorni di fila"
            avvisi.append(Problema(
                "avviso", f"{s.cognome} {s.nome} ({s.classe}ª)",
                f"le 2 ore di 1° strumento sono {come} ({GIORNI_LUNGHI[gg[0]]} e {GIORNI_LUNGHI[gg[1]]}): "
                f"con le ore dichiarate da {s.doc1} non c'era modo di distanziarle di più.",
                "2 ore di 1° strumento in giorni vicini"))

    def buchi_di(lezioni: list[Lezione]) -> list[tuple[int, int]]:
        per_giorno: dict[int, set[int]] = defaultdict(set)
        for l in lezioni:
            per_giorno[l.giorno].add(l.ora)
        out = []
        for g, ore in sorted(per_giorno.items()):
            for o in range(min(ore), max(ore) + 1):
                if o not in ore:
                    out.append((g, o))
        return out

    for s in dati.studenti:
        mie = orario.lezioni_di(s.id)
        dove = f"Studenti, riga {s.riga}"
        buchi = buchi_di(mie)
        if buchi:
            per_g = Counter(g for g, _ in buchi)
            testo = "; ".join(
                f"{n} buc{'o' if n == 1 else 'hi'} il {_giorno(g)} ({', '.join(ORE[o] for gg, o in buchi if gg == g)})"
                for g, n in sorted(per_g.items())
            )
            avvisi.append(Problema("avviso", dove, f"{_etichetta(s)}: {testo}"))
        lmc = next((l for l in mie if l.tipo == TIPO_LMC), None)
        if lmc is not None:
            staccate = [l for l in mie if l.tipo in (TIPO_STRUM1, TIPO_STRUM2) and l.giorno != lmc.giorno]
            if staccate:
                quali = ", ".join(f"{_NOME_TIPO[l.tipo]} il {_giorno(l.giorno)}" for l in staccate)
                avvisi.append(Problema("avviso", dove,
                                       f"{_etichetta(s)}: la musica da camera è il {_giorno(lmc.giorno)}, "
                                       f"ma {quali} (non si è riusciti ad attaccarle)"))
        rientri = len({l.giorno for l in mie})
        if rientri > par.max_rientri:
            avvisi.append(Problema("avviso", dove, f"{_etichetta(s, True)}: {rientri} rientri"))
        if s.trasporto is not None and s.trasporto.valido:
            senza_mezzo = [l for l in mie if l.ora < len(s.trasporto.minuti) and s.trasporto.minuti[l.ora] is None]
            if senza_mezzo:
                testo = ", ".join(f"{_giorno(l.giorno)} alle {ORE[l.ora]}" for l in senza_mezzo)
                avvisi.append(Problema("avviso", dove, f"{_etichetta(s)}: lezione {testo} ma dopo quell'ora "
                                                        "non risulta un mezzo per tornare a casa"))
            elif not _vicino(s, dati):
                tardive = [l for l in mie if l.ora >= N_ORE - 1 and s.trasporto.minuti[l.ora] is not None
                           and s.trasporto.minuti[l.ora] >= 90]
                if tardive:
                    testo = ", ".join(f"{_giorno(l.giorno)} alle {ORE[l.ora]} (a casa alle {s.trasporto.arrivi[l.ora] or '?'})"
                                      for l in tardive)
                    avvisi.append(Problema("avviso", dove, f"{_etichetta(s, True)}: lezione {testo}"))
        elif not _vicino(s, dati):
            soglia_ora = N_ORE - 1 if km_norm[s.id] < 0.5 else N_ORE - 2
            tardive = [l for l in mie if l.ora >= soglia_ora]
            if tardive:
                testo = ", ".join(f"lezione alle {ORE[l.ora]} il {_giorno(l.giorno)}" for l in tardive)
                avvisi.append(Problema("avviso", dove, f"{_etichetta(s, True)}: {testo}"))

    for d in dati.docenti:
        buchi = buchi_di(orario.lezioni_docente(d.nome))
        if buchi:
            per_g = Counter(g for g, _ in buchi)
            testo = "; ".join(
                f"{n} buc{'o' if n == 1 else 'hi'} il {_giorno(g)} ({', '.join(ORE[o] for gg, o in buchi if gg == g)})"
                for g, n in sorted(per_g.items())
            )
            avvisi.append(Problema("avviso", f"Docenti: {d.nome}", f"Docente {d.nome}: {testo}"))
    return avvisi


# ── Diagnosi quando non si trova una soluzione ───────────────────────────────

def _diagnosi(dati: DatiInput, unita: list[_Unita], stato_strict: int,
              progresso: Callable[[str], None]) -> list[Problema]:
    """Rilancia il modello lasciando fuori il minimo numero di lezioni e spiega quali e perché."""
    modello = _Modello(dati, unita, diagnosi=True)
    timeout = min(TIMEOUT_DIAGNOSI_S, max(5, dati.parametri.timeout_s))
    solver, stato = _risolvi(modello, timeout, progresso)
    problemi: list[Problema] = []

    if stato in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        fuori = [u for u in unita if solver.value(modello.manca[u.idx])]
        posizione = {u.idx: modello.fascia_di(solver, u) for u in unita}
        if fuori:
            testa = (f"Non è stato possibile collocare {len(fuori)} lezion{'e' if len(fuori) == 1 else 'i'}: "
                     "l'orario non viene prodotto. Qui sotto il dettaglio.")
            problemi.append(Problema("errore", "Calcolo", testa))
            for u, n in _raggruppa(fuori):
                problemi.append(_spiega_unita(dati, unita, u, posizione, n))
            return problemi
        # tutte collocate nel modello rilassato: la versione rigida era infattibile solo per il tempo/prove
        if stato_strict == cp_model.INFEASIBLE:
            problemi.append(Problema("errore", "Calcolo",
                                     "I vincoli si contraddicono tra loro ma non è stato possibile isolare la lezione responsabile. "
                                     "Controllare i vincoli 'Giorno unico', '1° strumento attaccato' e i giorni non disponibili."))
        else:
            problemi.append(Problema("errore", "Calcolo",
                                     f"Nessuna combinazione trovata entro {dati.parametri.timeout_s} secondi. "
                                     "Il problema è molto stretto: provare ad aumentare il tempo massimo nel foglio Parametri "
                                     "o ad allargare le disponibilità dei docenti più carichi."))
    else:
        problemi.append(Problema("errore", "Calcolo",
                                 "Nessuna combinazione trovata: le regole rigide non si riescono a rispettare tutte insieme "
                                 "e il programma non è riuscito a individuare una singola lezione responsabile."))
    problemi.extend(_controlli_euristici(dati, unita))
    return problemi


def _spiega_unita(dati: DatiInput, unita: list[_Unita], u: _Unita,
                  posizione: dict[int, int | None] | None, n: int = 1) -> Problema:
    """Un `Problema` in italiano per una lezione non collocabile, con il probabile motivo."""
    docente = dati.docente(u.docente)
    studenti = [dati.studente(i) for i in u.studenti]
    studenti = [s for s in studenti if s is not None]
    dove = f"Docenti: {u.docente}" if not studenti else f"Studenti, riga {studenti[0].riga}"
    chi = ", ".join(_etichetta(s) for s in studenti) if studenti else f"Docente {u.docente}"
    articolo = "l'" if n == 1 or u.tipo == TIPO_LMC else "le "
    frasi = [f"{chi}: impossibile collocare {articolo}{u.descrizione(dati, n)}."]

    vietati: set[int] = set()
    for s in studenti:
        vietati |= s.giorni_non_disp
    giorni_ok = [g for g in range(N_GIORNI) if g not in vietati]
    fasce_doc = [f for f, v in docente.disponibilita.items()
                 if (v == "X" or (u.tipo == TIPO_ACCOMP and v == "A")) and giorno_ora(f)[0] in giorni_ok]

    if not docente.disponibilita:
        frasi.append(f"Il docente {u.docente} non ha nessuna X nel foglio Docenti.")
    elif not fasce_doc:
        frasi.append(f"Il docente {u.docente} non ha nessuna fascia disponibile nei giorni in cui "
                     f"{'lo studente può venire' if len(studenti) == 1 else 'tutti i ragazzi del gruppo possono venire'} "
                     f"({', '.join(_giorno(g) for g in giorni_ok) or 'nessun giorno'}).")
    else:
        elenco_giorni = ", ".join(_giorno(g) for g in giorni_ok)
        if vietati:
            frasi.append(f"Il docente ha {len(fasce_doc)} fasce disponibili nei giorni ammessi ({elenco_giorni}).")
        else:
            frasi.append(f"Il docente ha {len(fasce_doc)} fasce disponibili.")
        if u.seguente is not None or u.seconda:
            coppie = [f for f in fasce_doc if giorno_ora(f)[1] < N_ORE - 1 and (f + 1) in fasce_doc]
            frasi.append(f"Servono due fasce consecutive dello stesso giorno: il docente ne ha {len(coppie)} coppie.")
        if posizione is not None:
            # come sono usate quelle fasce nel miglior orario trovato
            occ_doc = {posizione[v.idx] for v in unita if v.docente == u.docente and v.idx != u.idx}
            ids = set(u.studenti)
            occ_stud = {posizione[v.idx] for v in unita if v.idx != u.idx and ids & set(v.studenti)}
            n_doc = sum(1 for f in fasce_doc if f in occ_doc)
            n_stud = sum(1 for f in fasce_doc if f not in occ_doc and f in occ_stud)
            n_lib = len(fasce_doc) - n_doc - n_stud
            parti = []
            if n_doc:
                parti.append(f"{n_doc} occupate da altre lezioni del docente")
            if n_stud:
                parti.append(f"{n_stud} in cui {'lo studente ha' if len(studenti) == 1 else 'un ragazzo del gruppo ha'} già un'altra lezione")
            if parti:
                frasi.append("Nel miglior orario trovato sono " + " e ".join(parti) + ".")
            if n_lib > 0:
                frasi.append(f"Restano {n_lib} fasce libere ma non compatibili con gli altri vincoli "
                             f"({'giorno unico, ' if any(s.giorno_unico for s in studenti) else ''}"
                             f"{'1° strumento attaccato, ' if u.seguente is not None or u.seconda else ''}massimo rientri).")
    if docente.fasce_a and u.tipo != TIPO_ACCOMP:
        frasi.append(f"Attenzione: il docente ha {len(docente.fasce_a)} A fissate "
                     f"({', '.join(nome_fascia(f) for f in docente.fasce_a)}) e le lezioni della stessa giornata "
                     "devono stare prima dell'accompagnamento: le X dopo una A non sono utilizzabili per le lezioni.")
    for s in studenti:
        if s.giorno_unico:
            docs = [dati.docente(n) for n in (s.doc1, s.doc2) if n]
            g_lmc = dati.gruppo_di(s.id)
            if g_lmc:
                docs.append(dati.docente(g_lmc.docente))
            docs = [d for d in docs if d is not None]
            comuni = [g for g in range(N_GIORNI) if g not in s.giorni_non_disp
                      and all(any(giorno_ora(f)[0] == g for f in d.fasce_x) for d in docs)]
            frasi.append(f"{_etichetta(s)} deve venire un solo giorno: i giorni in cui tutti i suoi docenti "
                         f"sono disponibili sono {', '.join(_giorno(g) for g in comuni) or 'nessuno'}.")
    if u.tipo == TIPO_LMC and len(studenti) > 1:
        frasi.append(f"Il gruppo ha {len(studenti)} ragazzi che devono essere liberi tutti nella stessa fascia.")
    frasi.append(f"Suggerimento: aggiungere qualche X al docente {u.docente}"
                 + (" nei giorni indicati" if vietati else "")
                 + (", oppure allentare i vincoli dello studente (giorno unico, 1° strumento attaccato, giorni non disponibili)."
                    if studenti else "."))
    return Problema("errore", dove, " ".join(frasi))


def _controlli_euristici(dati: DatiInput, unita: list[_Unita]) -> list[Problema]:
    """Indizi generici sui punti più stretti del problema (usati quando la diagnosi precisa fallisce)."""
    problemi: list[Problema] = []
    carico: Counter = Counter(u.docente for u in unita)
    for d in dati.docenti:
        n = carico.get(d.nome, 0)
        disp = len(d.disponibilita)
        if n and disp and n / disp >= 0.9:
            problemi.append(Problema("avviso", f"Docenti: {d.nome}",
                                     f"Il docente {d.nome} ha {n} ore da collocare su {disp} fasce disponibili: margine quasi nullo."))
    for u in unita:
        if u.tipo == TIPO_ACCOMP:
            continue
        if len(u.fasce) <= 2:
            chi = ", ".join(_etichetta(s) for s in (dati.studente(i) for i in u.studenti) if s)
            problemi.append(Problema("avviso", f"Docenti: {u.docente}",
                                     f"{chi}: l'{u.descrizione(dati)} ha solo {len(u.fasce)} fasce possibili "
                                     f"({', '.join(nome_fascia(f) for f in u.fasce)})."))
    for s in dati.studenti:
        if not s.giorno_unico:
            continue
        docs = [dati.docente(n) for n in (s.doc1, s.doc2) if n]
        g_lmc = dati.gruppo_di(s.id)
        if g_lmc:
            docs.append(dati.docente(g_lmc.docente))
        docs = [d for d in docs if d is not None]
        comuni = [g for g in range(N_GIORNI) if g not in s.giorni_non_disp
                  and all(any(giorno_ora(f)[0] == g for f in d.fasce_x) for d in docs)]
        if len(comuni) <= 1:
            problemi.append(Problema("avviso", f"Studenti, riga {s.riga}",
                                     f"{_etichetta(s)}: giorno unico con un solo giorno possibile "
                                     f"({', '.join(_giorno(g) for g in comuni) or 'nessuno'})."))
    return problemi


__all__ = ["calcola"]
