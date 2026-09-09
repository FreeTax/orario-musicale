"""Prova del motore CP-SAT sui dati di prova (`input_prova.xlsx`).

Esegue: calcolo completo con statistiche e verifiche, un caso impossibile (deve sollevare
ProblemiError con messaggi leggibili) e un ricalcolo a modifiche minime con `precedente`.

Uso:  .venv/bin/python test_motore.py [timeout_s]
"""

from __future__ import annotations

import copy
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

from orario.controlli import controlla
from orario.costanti import N_ORE, ORE, TIPO_ACCOMP, TIPO_LMC, TIPO_STRUM1, TIPO_STRUM2
from orario.lettura import leggi_dati
from orario.modello import DatiInput, Orario, ProblemiError
from orario.motore import calcola

PERCORSO = Path(__file__).parent / "input_prova.xlsx"
TIMEOUT = int(sys.argv[1]) if len(sys.argv) > 1 else 60


def progresso(msg: str) -> None:
    print(f"    [{time.strftime('%H:%M:%S')}] {msg}")


def carica() -> DatiInput:
    if not PERCORSO.exists():
        sys.exit("Manca input_prova.xlsx: eseguire  .venv/bin/python dati_prova.py")
    dati = leggi_dati(PERCORSO)
    avvisi = controlla(dati)
    dati.parametri.timeout_s = TIMEOUT
    print(f"Dati: {len(dati.studenti)} studenti, {len(dati.docenti)} docenti, {len(dati.gruppi)} gruppi LMC, "
          f"{len(avvisi)} avvisi dai controlli")
    return dati


def buchi_docente(orario: Orario, nome: str) -> int:
    per_giorno: dict[int, list[int]] = defaultdict(list)
    for l in orario.lezioni_docente(nome):
        per_giorno[l.giorno].append(l.ora)
    return sum((max(o) - min(o) + 1) - len(o) for o in per_giorno.values())


def statistiche(orario: Orario) -> None:
    dati = orario.dati
    print(f"  Stato: {orario.stato}   tempo: {orario.secondi} s   lezioni: {len(orario.lezioni)}")
    tot_buchi = sum(orario.buchi_di(s.id) for s in dati.studenti)
    print(f"  Buchi studenti totali: {tot_buchi}")
    rientri = Counter(orario.rientri_di(s.id) for s in dati.studenti)
    print("  Rientri: " + ", ".join(f"{n} giorno/i → {rientri[n]} studenti" for n in sorted(rientri)))
    fasce_km = [(0, 5), (5, 15), (15, 25), (25, 999)]
    for a, b in fasce_km:
        gruppo = [s for s in dati.studenti if s.km is not None and a <= s.km < b]
        if not gruppo:
            continue
        ore = [l.ora for s in gruppo for l in orario.lezioni_di(s.id)]
        media = sum(ore) / len(ore)
        print(f"  km {a:>2}-{b if b < 999 else '…':<3}: {len(gruppo):>3} studenti, ora media {media:.2f} "
              f"(0 = {ORE[0]}, 3 = {ORE[3]})")
    tot_bd = 0
    con_buchi = []
    for d in dati.docenti:
        b = buchi_docente(orario, d.nome)
        tot_bd += b
        if b:
            con_buchi.append(f"{d.nome} ({b})")
    print(f"  Buchi docenti totali: {tot_bd}" + (f" → {', '.join(con_buchi)}" if con_buchi else ""))
    print(f"  Avvisi: {len(orario.avvisi)}")
    for a in orario.avvisi[:15]:
        print(f"    - {a.messaggio}")


def verifica(orario: Orario) -> None:
    dati = orario.dati
    docenti = {d.nome: d for d in dati.docenti}
    # ore per studente
    for s in dati.studenti:
        mie = orario.lezioni_di(s.id)
        tipi = Counter(l.tipo for l in mie)
        assert (tipi[TIPO_STRUM1], tipi[TIPO_STRUM2], tipi[TIPO_LMC]) == s.ore, f"{s.id}: ore {tipi} attese {s.ore}"
        assert len({l.fascia for l in mie}) == len(mie), f"{s.id}: due lezioni nella stessa fascia"
        for l in mie:
            assert l.giorno not in s.giorni_non_disp, f"{s.id}: lezione in giorno vietato"
        if s.giorno_unico:
            assert orario.rientri_di(s.id) == 1, f"{s.id}: giorno unico non rispettato"
        if s.ore_consecutive and s.ore[0] == 2:
            a, b = sorted((l for l in mie if l.tipo == TIPO_STRUM1), key=lambda l: l.fascia)
            assert a.giorno == b.giorno and b.ora == a.ora + 1, f"{s.id}: ore non consecutive"
            assert b.due_ore and not a.due_ore, f"{s.id}: due_ore sbagliato"
        assert orario.rientri_di(s.id) <= dati.parametri.max_rientri_vicini
    # docenti
    conta = Counter((l.docente, l.fascia) for l in orario.lezioni)
    assert all(n == 1 for n in conta.values()), "docente doppio in una fascia"
    for l in orario.lezioni:
        v = docenti[l.docente].disponibilita.get(l.fascia)
        assert v in (("X", "A") if l.tipo == TIPO_ACCOMP else ("X",)), f"{l.docente} fascia {l.fascia} disp={v}"
    for d in dati.docenti:
        acc = [l for l in orario.lezioni if l.docente == d.nome and l.tipo == TIPO_ACCOMP]
        assert len(acc) == d.ore_accomp, f"{d.nome}: {len(acc)} ACCOMP invece di {d.ore_accomp}"
        for f in d.fasce_a:
            assert sum(1 for l in acc if l.fascia == f) == 1, f"{d.nome}: A in fascia {f} non rispettata"
        lez = [l for l in orario.lezioni if l.docente == d.nome and l.tipo != TIPO_ACCOMP]
        for g in range(5):
            oa = [l.ora for l in acc if l.giorno == g]
            ol = [l.ora for l in lez if l.giorno == g]
            assert not (oa and ol) or min(oa) > max(ol), f"{d.nome}: accompagnamento prima di una lezione"
    # LMC insieme
    for g in dati.gruppi:
        lg = [l for l in orario.lezioni if l.tipo == TIPO_LMC and l.gruppo == g.numero]
        assert len(lg) == 1 and set(lg[0].studenti) == set(g.studenti) and lg[0].docente == g.docente
    print("  Verifiche: OK")


def caso_impossibile(dati: DatiInput) -> None:
    print("\n=== Caso impossibile 1: docente senza nessuna X ===")
    d2 = copy.deepcopy(dati)
    doc = next(d for d in d2.docenti if d.nome == "Bini")
    doc.disponibilita.clear()
    try:
        calcola(d2, progresso)
        raise AssertionError("doveva sollevare ProblemiError")
    except ProblemiError as e:
        print(f"  ProblemiError con {len(e.problemi)} problemi. Primi 3:")
        for p in e.problemi[:3]:
            print(f"    - {p}")

    print("\n=== Caso impossibile 2: due studenti che vogliono la stessa unica fascia (i controlli non lo vedono) ===")
    d3 = copy.deepcopy(dati)
    doc = next(d for d in d3.docenti if len(d.fasce_x) >= 18)
    stud = [s for s in d3.studenti if s.doc2 == doc.nome and s.ore[1] == 1][:2]
    if len(stud) < 2:
        stud = [s for s in d3.studenti if s.doc1 == doc.nome][:2]
    f_unica = doc.fasce_x[0]
    giorno_unico = f_unica // N_ORE
    # il docente resta disponibile in quel giorno solo nella fascia f_unica
    doc.disponibilita = {f: v for f, v in doc.disponibilita.items() if f // N_ORE != giorno_unico or f == f_unica}
    for s in stud:
        s.giorni_non_disp = {g for g in range(5) if g != giorno_unico}
    controlla(d3)  # non deve fermarsi: i controlli elementari passano
    d3.parametri.timeout_s = TIMEOUT
    try:
        calcola(d3, progresso)
        raise AssertionError("doveva sollevare ProblemiError")
    except ProblemiError as e:
        print(f"  ProblemiError con {len(e.problemi)} problemi:")
        for p in e.problemi:
            print(f"    - {p}")


def caso_ricalcolo(dati: DatiInput, orario1: Orario) -> None:
    print("\n=== Ricalcolo a modifiche minime: tolgo una X occupata a un docente ===")
    d2 = copy.deepcopy(dati)
    # un docente con margine e una sua lezione con studenti
    lez = next(l for l in orario1.lezioni if l.tipo in (TIPO_STRUM1, TIPO_STRUM2)
               and len(d2.docente(l.docente).fasce_x) >= 18)
    doc = d2.docente(lez.docente)
    del doc.disponibilita[lez.fascia]
    print(f"  Tolta la X a {doc.nome} nella fascia {lez.fascia} (occupata da {lez.studenti[0]})")
    controlla(d2)
    d2.parametri.timeout_s = TIMEOUT
    orario2 = calcola(d2, progresso, precedente=orario1.lezioni)
    statistiche(orario2)
    verifica(orario2)
    prec = {(l.docente, l.tipo, frozenset(l.studenti)): [] for l in orario1.lezioni}
    for l in orario1.lezioni:
        prec[(l.docente, l.tipo, frozenset(l.studenti))].append(l.fascia)
    spostate = 0
    for k, fasce in prec.items():
        nuove = sorted(l.fascia for l in orario2.lezioni if (l.docente, l.tipo, frozenset(l.studenti)) == k)
        spostate += sum(1 for a, b in zip(sorted(fasce), nuove) if a != b)
    print(f"  Lezioni spostate rispetto all'orario 1: {spostate}")
    print("  Avvisi di spostamento:")
    for a in orario2.avvisi:
        if a.dove == "Ricalcolo" or "spostat" in a.messaggio:
            print(f"    - {a.messaggio}")
    assert spostate <= 12, f"troppe lezioni spostate: {spostate}"


def main() -> None:
    dati = carica()
    print(f"\n=== Calcolo completo (timeout {TIMEOUT} s) ===")
    orario = calcola(dati, progresso)
    statistiche(orario)
    verifica(orario)
    caso_impossibile(dati)
    caso_ricalcolo(dati, orario)
    print("\nTutto OK.")


if __name__ == "__main__":
    main()
