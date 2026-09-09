"""Crea `input_prova.xlsx`: una copia di `input_orario.xlsx` completata con dati INVENTATI
(docenti della 1ª, km della 1ª, disponibilità dei docenti, gruppi LMC, LMI, eccezioni),
per poter provare il programma prima che arrivi il dataset vero.

Uso:  .venv/bin/python dati_prova.py
"""

from __future__ import annotations

import random
import shutil
from collections import Counter, defaultdict
from pathlib import Path

from orario.costanti import FASCE, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_LMI, FOGLIO_PARAMETRI, FOGLIO_STUDENTI, N_ORE
from orario.lettura import aggiorna_struttura, leggi_tabelle, salva_tabelle

ROOT = Path(__file__).parent
SORGENTE = ROOT / "input_orario.xlsx"
DESTINAZIONE = ROOT / "input_prova.xlsx"
MAX_ORE_DOCENTE = 17  # oltre questo carico, per la prova, si inventa un secondo docente
random.seed(2026)

SINONIMI = {"C.BASSO": "CONTRABBASSO", "PERCUSSIONE": "PERCUSSIONI", "V.CELLO": "VIOLONCELLO"}


def strum(s: str) -> str:
    s = (s or "").upper().strip()
    return SINONIMI.get(s, s)


def main() -> None:
    shutil.copy(SORGENTE, DESTINAZIONE)
    tab = leggi_tabelle(DESTINAZIONE)
    aggiorna_struttura(tab)   # il file di partenza può essere di una versione precedente
    ts, td, tg, tl = tab[FOGLIO_STUDENTI], tab[FOGLIO_DOCENTI], tab[FOGLIO_GRUPPI], tab[FOGLIO_LMI]
    c = {n: ts.colonna(n) for n in ["Classe", "Cognome", "Nome", "KM", "Strumento 1", "Docente 1",
                                    "Strumento 2", "Docente 2", "Ore consecutive", "Giorno unico", "Giorni NON"]}

    # ── docenti per strumento e carico attuale (classi 2ª-5ª) ──
    doc_per_strum: dict[str, set[str]] = defaultdict(set)
    ore: Counter = Counter()
    for r in ts.righe:
        cl = int(r[c["Classe"]])
        h1 = 2 if cl in (1, 2, 5) else 1
        if r[c["Docente 1"]]:
            doc_per_strum[strum(r[c["Strumento 1"]])].add(r[c["Docente 1"]])
            ore[r[c["Docente 1"]]] += h1
        if cl != 5 and r[c["Docente 2"]]:
            doc_per_strum[strum(r[c["Strumento 2"]])].add(r[c["Docente 2"]])
            ore[r[c["Docente 2"]]] += 1

    nuovi_docenti: dict[str, str] = {}  # nome → strumento
    ore["Paccosi"] += 5  # ore di accompagnamento (prova), contano nel carico

    def scegli_docente(strumento: str, h: int) -> str:
        cand = sorted(doc_per_strum.get(strumento, set()), key=lambda d: ore[d])
        if cand and ore[cand[0]] + h <= MAX_ORE_DOCENTE:
            ore[cand[0]] += h
            return cand[0]
        nome = f"{strumento.title()} 2 (prova)"
        if nome not in nuovi_docenti:
            nuovi_docenti[nome] = strumento
            doc_per_strum[strumento].add(nome)
        ore[nome] += h
        return nome

    # ── completa la 1ª: docenti e km; eccezioni sparse ──
    for r in ts.righe:
        cl = int(r[c["Classe"]])
        if cl == 1:
            r[c["Docente 1"]] = r[c["Docente 1"]] or scegli_docente(strum(r[c["Strumento 1"]]), 2)
            r[c["Docente 2"]] = r[c["Docente 2"]] or scegli_docente(strum(r[c["Strumento 2"]]), 1)
            r[c["KM"]] = r[c["KM"]] or str(round(random.choice([0.8, 2.5, 4, 7, 12, 18, 25, 31]) + random.random(), 1))
        x = random.random()
        if cl in (1, 2, 5) and x < 0.10:
            r[c["Ore consecutive"]] = "SI"
        elif x < 0.13:
            r[c["Giorno unico"]] = "SI"
        elif x < 0.17:
            r[c["Giorni NON"]] = random.choice(["Mar", "Gio", "Lun, Ven"])

    # ── gruppi LMC (classi 3ª-5ª), docenti LMC come nell'anno vecchio ──
    docenti_lmc = ["Di Bonaventura", "Donnini", "Corsini", "Pacini M", "Bini", "Simonelli", "Lala", "Lissoni"]
    triennio = [r for r in ts.righe if int(r[c["Classe"]]) in (3, 4, 5)]
    random.shuffle(triennio)
    tg.righe = []
    i, n = 0, 1
    while i < len(triennio):
        k = min(random.choice([2, 2, 3, 3, 4, 5]), len(triennio) - i)
        if len(triennio) - i - k == 1:
            k += 1
        membri = triennio[i:i + k]
        i += k
        doc = min(docenti_lmc, key=lambda d: ore[d])
        ore[doc] += 1
        nomi = [f"{m[c['Cognome']].title()}" for m in membri]
        # cognomi doppi nel triennio → aggiungi il nome
        for j, m in enumerate(membri):
            if sum(1 for r in ts.righe if r[c["Cognome"]] == m[c["Cognome"]] and int(r[c["Classe"]]) in (3, 4, 5)) > 1:
                nomi[j] = f"{m[c['Cognome']].title()} {m[c['Nome']].title()}"
        tg.righe.append([str(n), doc] + nomi + [""] * (5 - len(nomi)) + [""])
        n += 1

    # ── docenti: nuovi (prova), ore accompagnamento, disponibilità ──
    cd = {n: td.colonna(n) for n in ["Docente", "Strumento", "Aula Lun", "Ore accomp"]}
    idx_f = [td.colonna(f) for f in FASCE]
    for nome, s in nuovi_docenti.items():
        riga = [""] * len(td.intestazione)
        riga[cd["Docente"]], riga[cd["Strumento"]], riga[cd["Ore accomp"]] = nome, s, "0"
        for k in range(5):
            riga[cd["Aula Lun"] + k] = "M 12"
        td.righe.append(riga)
    for r in td.righe:
        nome = r[cd["Docente"]]
        r[cd["Ore accomp"]] = "5" if nome == "Paccosi" else "0"
        tot = ore[nome] + int(r[cd["Ore accomp"]])
        giorni = list(range(5))
        random.shuffle(giorni)
        n_giorni = 5 if tot > 15 else 4 if tot > 10 else 3 if tot > 5 else 2
        for f in idx_f:
            r[f] = ""
        for g in giorni[:n_giorni]:
            for o in range(N_ORE):
                if random.random() < 0.92:
                    r[idx_f[g * N_ORE + o]] = "X"
        # garantisce margine: se le X sono poche, riempi
        while sum(1 for f in idx_f if r[f]) < tot + 3 and sum(1 for f in idx_f if r[f]) < 20:
            r[idx_f[random.randrange(20)]] = "X"
        if nome == "Paccosi":  # un'ora di accompagnamento fissata a mano
            for f in idx_f:
                if r[f] == "X":
                    r[f] = "A"
                    break

    # ── LMI (solo da ricopiare) ──
    tl.righe = [
        ["LMI PERCUSSIONI", "1ª", "Mazzei", "Aula Perc 1, 1° piano", "Martedì 4ª-5ª ora", "Del Rosso, Stolfi, Zanieri", ""],
        ["LMI FIATI", "1ª", "Donnini", "Aula Sax, 1° piano", "Martedì 4ª-5ª ora", "Boldrini, Orsucci, Pratesi, Paganelli, Ceccarelli, Grippaldi, Li", ""],
        ["LMI ARCHI 2ª-3ª", "2ª, 3ª", "Di Bonaventura", "Aula 5ALM, 1° piano", "Giovedì 1ª-2ª ora", "Barneschi, Del Pero, Greco, Ioanna, Vaiani", ""],
        ["LMI CHITARRE 4ª-5ª", "4ª, 5ª", "Bini", "Aula Chitarra, 2° piano", "Venerdì 3ª-4ª ora", "Manganiello, Ottaviani, Valentini", ""],
    ]

    # indirizzo della scuola (di prova) nel foglio Parametri
    tp = tab[FOGLIO_PARAMETRI]
    for r in tp.righe:
        if "scuola" in r[0].lower():
            r[1] = "Corso Gramsci 148, Pistoia"

    salva_tabelle(DESTINAZIONE, tab)
    print(f"Creato {DESTINAZIONE.name}: {len(ts.righe)} studenti, {len(td.righe)} docenti, {len(tg.righe)} gruppi LMC.")
    if nuovi_docenti:
        print("Docenti inventati per la prova:", ", ".join(nuovi_docenti))
    print("Carico per docente (ore lezione + LMC):")
    for d, h in sorted(ore.items(), key=lambda x: -x[1]):
        print(f"  {d:22s} {h}")


if __name__ == "__main__":
    main()
