"""Prova dell'export: costruisce un orario finto (assegnazione greedy, senza preferenze)
partendo dai dati veri di `input_prova.xlsx` e produce i file in `output_prova/`.

Uso:  .venv/bin/python test_export.py
Se `orario/motore.py` con `calcola(dati)` è disponibile, prova anche con l'orario vero
in `output_prova_motore/`.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

from orario.controlli import controlla
from orario.costanti import N_FASCE, N_ORE, TIPO_ACCOMP, TIPO_LMC, TIPO_STRUM1, TIPO_STRUM2
from orario.export_excel import esporta_tutto
from orario.lettura import leggi_dati
from orario.modello import DatiInput, Lezione, Orario, Problema

ROOT = Path(__file__).parent
INPUT = ROOT / "input_prova.xlsx"


def orario_greedy(dati: DatiInput) -> Orario:
    """Assegnazione ingenua: prima fascia in cui docente e studenti sono liberi."""
    t0 = time.perf_counter()
    occupato_doc: dict[str, set[int]] = {d.nome: set() for d in dati.docenti}
    occupato_stud: dict[str, set[int]] = {s.id: set() for s in dati.studenti}
    lezioni: list[Lezione] = []
    avvisi: list[Problema] = []

    def libero(doc: str, f: int, ids: list[str]) -> bool:
        d = dati.docente(doc)
        return (d is not None and d.disponibilita.get(f) == "X" and f not in occupato_doc[doc]
                and all(f not in occupato_stud[i] and f // N_ORE not in dati.studente(i).giorni_non_disp for i in ids))

    def assegna(doc: str, f: int, tipo: str, ids: list[str], **kw) -> None:
        lezioni.append(Lezione(doc, f, tipo, list(ids), **kw))
        occupato_doc[doc].add(f)
        for i in ids:
            occupato_stud[i].add(f)

    # gruppi LMC prima (sono i più vincolati)
    for g in dati.gruppi:
        for f in range(N_FASCE):
            if libero(g.docente, f, g.studenti):
                assegna(g.docente, f, TIPO_LMC, g.studenti, gruppo=g.numero)
                break
        else:
            avvisi.append(Problema("avviso", f"Gruppo LMC {g.numero}", "Non collocato dalla prova greedy."))

    # strumenti, studenti più lontani prima
    for s in sorted(dati.studenti, key=lambda x: -(x.km or 0)):
        h1, h2, _ = s.ore
        if h1 and s.ore_consecutive and h1 == 2:
            for f in range(N_FASCE):
                if f % N_ORE < N_ORE - 1 and libero(s.doc1, f, [s.id]) and libero(s.doc1, f + 1, [s.id]):
                    assegna(s.doc1, f, TIPO_STRUM1, [s.id])
                    assegna(s.doc1, f + 1, TIPO_STRUM1, [s.id], due_ore=True)
                    break
            else:
                avvisi.append(Problema("avviso", f"{s.cognome} {s.nome}", "Ore consecutive non collocate dalla prova greedy."))
            h1 = 0
        for doc, ore, tipo in ((s.doc1, h1, TIPO_STRUM1), (s.doc2, h2, TIPO_STRUM2)):
            for _ in range(ore):
                # preferisce i giorni in cui lo studente è già presente
                giorni_presenti = {f // N_ORE for f in occupato_stud[s.id]}
                candidati = sorted(range(N_FASCE), key=lambda f: (f // N_ORE not in giorni_presenti, f))
                for f in candidati:
                    if libero(doc, f, [s.id]):
                        assegna(doc, f, tipo, [s.id])
                        break
                else:
                    avvisi.append(Problema("avviso", f"{s.cognome} {s.nome} ({s.classe}ª)",
                                           f"Ora di {tipo} con {doc} non collocata dalla prova greedy."))

    # accompagnamento: fasce A, altrimenti X libere
    for d in dati.docenti:
        n = 0
        for f in d.fasce_a + [f for f in d.fasce_x if f not in occupato_doc[d.nome]]:
            if n >= d.ore_accomp:
                break
            if f not in occupato_doc[d.nome]:
                assegna(d.nome, f, TIPO_ACCOMP, [])
                n += 1

    return Orario(lezioni=lezioni, dati=dati, stato="PROVA GREEDY", secondi=time.perf_counter() - t0, avvisi=avvisi)


def main() -> None:
    if not INPUT.exists():
        sys.exit("Manca input_prova.xlsx: eseguire prima  .venv/bin/python dati_prova.py")
    dati = leggi_dati(INPUT)
    controlla(dati)
    orario = orario_greedy(dati)
    print(f"Orario greedy: {len(orario.lezioni)} lezioni, {len(orario.avvisi)} avvisi")
    for p in esporta_tutto(orario, ROOT / "output_prova"):
        print("  ->", p)

    try:
        from orario.motore import calcola  # type: ignore
    except ImportError:
        print("orario/motore.py non disponibile: saltata la prova con l'orario vero.")
        return
    print("Calcolo con il motore vero...")
    try:
        vero = calcola(dati)
    except Exception as e:  # il motore può essere ancora in lavorazione
        print(f"Il motore ha fallito ({type(e).__name__}: {e}): saltata la prova con l'orario vero.")
        return
    print(f"Orario motore: {len(vero.lezioni)} lezioni, stato {vero.stato}, {vero.secondi:.1f} s")
    for p in esporta_tutto(vero, ROOT / "output_prova_motore"):
        print("  ->", p)


if __name__ == "__main__":
    main()
