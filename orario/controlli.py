"""Controlli preliminari sui dati, prima di avviare il calcolo.

Ogni problema è spiegato in italiano semplice, con il punto del file in cui intervenire.
Gli "errori" fermano il calcolo; gli "avvisi" vengono mostrati ma non bloccano.
Come effetto collaterale risolve i cognomi dei gruppi LMC in id studente.
"""

from __future__ import annotations

import difflib
from collections import Counter, defaultdict

from .costanti import (
    CLASSI_LMC, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_LMI, FOGLIO_STUDENTI,
    MAX_GRUPPO_LMC, MIN_GRUPPO_LMC, N_FASCE, N_ORE, ORE_PER_CLASSE, nome_fascia,
)
from .modello import DatiInput, Problema, ProblemiError, Studente


def _norm(s: str) -> str:
    return " ".join((s or "").upper().replace("’", "'").split())


def risolvi_studente(testo: str, studenti: list[Studente]) -> tuple[Studente | None, str]:
    """Cerca uno studente a partire da 'Cognome', 'Cognome Nome' o 'Cognome N.'.

    Ritorna (studente, "") se unico, (None, motivo) altrimenti.
    """
    t = _norm(testo)
    if not t:
        return None, "nome vuoto"
    esatti = [s for s in studenti if s.id == t]
    if len(esatti) == 1:
        return esatti[0], ""
    per_cognome = [s for s in studenti if _norm(s.cognome) == t]
    if len(per_cognome) == 1:
        return per_cognome[0], ""
    if len(per_cognome) > 1:
        opzioni = ", ".join(f"{s.cognome} {s.nome} ({s.classe}ª)" for s in per_cognome)
        return None, f"ci sono più studenti con questo cognome ({opzioni}): scrivere anche il nome"
    # "Cognome N." oppure "Cognome Nome" con cognome composto
    parti = t.split(" ")
    for k in range(len(parti) - 1, 0, -1):
        cog, resto = " ".join(parti[:k]), " ".join(parti[k:]).rstrip(".")
        cand = [s for s in studenti if _norm(s.cognome) == cog and _norm(s.nome).startswith(resto)]
        if len(cand) == 1:
            return cand[0], ""
        if len(cand) > 1:
            return None, "più studenti corrispondono: scrivere il nome completo"
    # cognome che inizia con il testo (es. 'Degl'Innocenti' vs 'DEGL’INNOCENTI')
    cand = [s for s in studenti if _norm(s.cognome).replace("'", "") == t.replace("'", "")]
    if len(cand) == 1:
        return cand[0], ""
    # scritto con qualche lettera diversa: si prende il più somigliante, dicendolo
    per_id = {s.id: s for s in studenti}
    simili = difflib.get_close_matches(t, list(per_id), n=1, cutoff=0.86)
    if not simili:
        senza = {_norm(s.cognome): s for s in studenti}
        vicini = difflib.get_close_matches(t.split(" ")[0], list(senza), n=1, cutoff=0.86)
        if vicini and len({x for x in studenti if _norm(x.cognome) == vicini[0]}) == 1:
            s = senza[vicini[0]]
            return s, f"scritto «{testo}», inteso come «{s.cognome.title()} {s.nome.title()}»"
    else:
        s = per_id[simili[0]]
        return s, f"scritto «{testo}», inteso come «{s.cognome.title()} {s.nome.title()}»"
    return None, "non trovato nel foglio Studenti (controllare l'ortografia)"


def _ore_richieste(dati: DatiInput) -> Counter:
    """Ore pomeridiane che servono a ogni docente: lezioni individuali più musica da camera."""
    ore: Counter = Counter()
    for s in dati.studenti:
        h1, h2, _ = s.ore
        if s.doc1:
            ore[s.doc1] += h1
        if s.doc2 and h2:
            ore[s.doc2] += h2
    for g in dati.gruppi:
        if g.docente:
            ore[g.docente] += 1
    return ore


def apri_fasce_mancanti(dati: DatiInput) -> list[Problema]:
    """Se a un docente non bastano le ore dichiarate, ne apre quante ne servono e lo segnala.

    Meglio un orario da rivedere con il docente che nessun orario. Le fasce si scelgono
    dove disturbano meno: prima nei giorni in cui il docente c'è già, e nelle ore più presto.
    """
    segnalazioni: list[Problema] = []
    richieste = _ore_richieste(dati)
    for d in dati.docenti:
        servono = richieste.get(d.nome, 0) + d.ore_accomp
        disponibili = len(d.disponibilita)
        mancanti = servono - disponibili
        if mancanti <= 0 or servono > N_FASCE:
            continue
        giorni_gia = {f // N_ORE for f in d.disponibilita}
        libere = [f for f in range(N_FASCE) if f not in d.disponibilita]
        # ordine di scelta: giorno già usato, poi ora più presto, poi vicinanza a una fascia già libera
        libere.sort(key=lambda f: (0 if f // N_ORE in giorni_gia else 1,
                                   0 if (f - 1 in d.disponibilita or f + 1 in d.disponibilita) else 1,
                                   f % N_ORE, f))
        aggiunte = libere[:mancanti]
        for f in aggiunte:
            d.disponibilita[f] = "X"
        elenco = ", ".join(nome_fascia(f).replace("-", " alle ").replace(":", ".") for f in aggiunte)
        segnalazioni.append(Problema(
            "avviso", f"{FOGLIO_DOCENTI}: {d.nome}", categoria="ore aperte d'ufficio a un docente",
            messaggio=
            f"Servivano {servono} ore ma erano dichiarate disponibili solo {disponibili} fasce. "
            f"Ne ho aperte {len(aggiunte)} fra quelle segnate come non disponibili, per poter calcolare "
            f"l'orario: {elenco}. Da concordare con il docente."))
    return segnalazioni


def controlla(dati: DatiInput) -> list[Problema]:
    """Esegue tutti i controlli. Solleva ProblemiError se ci sono errori; ritorna gli avvisi."""
    problemi: list[Problema] = list(dati.avvisi)   # segnalazioni già emerse leggendo il file
    err = lambda dove, msg: problemi.append(Problema("errore", dove, msg))  # noqa: E731
    avv = lambda dove, msg, cat="": problemi.append(Problema("avviso", dove, msg, cat))  # noqa: E731

    nomi_docenti = [d.nome for d in dati.docenti]
    doppi = [n for n, c in Counter(nomi_docenti).items() if c > 1]
    for n in doppi:
        err(f"{FOGLIO_DOCENTI}: {n}", "Docente ripetuto due volte: tenere una sola riga.")
    docenti = {d.nome: d for d in dati.docenti}

    # ── Studenti ──
    ids = Counter(s.id for s in dati.studenti)
    for s in dati.studenti:
        dove = f"{FOGLIO_STUDENTI}, riga {s.riga} ({s.cognome} {s.nome}, {s.classe}ª)"
        if ids[s.id] > 1:
            err(dove, "Studente presente due volte con stesso cognome e nome.")
        h1, h2, _ = s.ore
        if h1 and not s.doc1:
            err(dove, "Manca il docente del 1° strumento.")
        elif h1 and s.doc1 not in docenti:
            err(dove, f"Il docente '{s.doc1}' del 1° strumento non è nel foglio {FOGLIO_DOCENTI}.")
        if h2 and not s.doc2:
            err(dove, "Manca il docente del 2° strumento.")
        elif h2 and s.doc2 not in docenti:
            err(dove, f"Il docente '{s.doc2}' del 2° strumento non è nel foglio {FOGLIO_DOCENTI}.")
        if s.km is None and s.trasporto is None:
            avv(dove, "Né KM né tempi dei mezzi: lo studente verrà trattato come se abitasse vicino alla scuola.",
                "ragazzi senza distanza né tempi dei mezzi")
        elif s.trasporto is not None and "centro del comune" in s.trasporto.esito:
            avv(dove, f"L'indirizzo «{s.indirizzo_completo}» non è sulle mappe: i tempi dei mezzi sono calcolati "
                      "dal centro del comune, quindi approssimati.",
                "indirizzi non trovati: tempi presi dal centro del comune")
        if s.ore_consecutive and h1 != 2:
            avv(dove, "Ore consecutive = SI ma la classe ha una sola ora di 1° strumento: ignorato.",
                "richieste di ore consecutive ignorate")
        if len(s.giorni_non_disp) >= 5:
            err(dove, "Lo studente non è disponibile in nessun giorno.")
        libere = [f for f in range(N_FASCE) if s.libero(f)]
        servono = sum(s.ore)
        if len(libere) < servono:
            err(dove, f"Fra giorni non disponibili e impegni restano {len(libere)} ore libere, "
                      f"ma ne servono {servono}.")
        elif s.fasce_non_disp and len(libere) <= servono + 1:
            avv(dove, f"Con gli impegni segnati restano solo {len(libere)} ore libere per {servono} lezioni: "
                      "poco margine.")

    # ── Gruppi LMC ──
    membri: dict[str, list[int]] = defaultdict(list)
    for g in dati.gruppi:
        dove = f"{FOGLIO_GRUPPI}, gruppo {g.numero} (riga {g.riga})"
        if not g.docente:
            err(dove, "Manca il docente del gruppo.")
        elif g.docente not in docenti:
            err(dove, f"Il docente '{g.docente}' non è nel foglio {FOGLIO_DOCENTI}.")
        risolti: list[str] = []
        for testo in g.studenti_raw:
            s, motivo = risolvi_studente(testo, dati.studenti)
            if s is None:
                err(dove, f"Studente '{testo}': {motivo}.")
                continue
            if motivo:
                avv(dove, motivo + ": conviene correggere il nome nel foglio.",
                    "nomi riconosciuti per somiglianza")
            if s.classe not in CLASSI_LMC:
                err(dove, f"{s.cognome} {s.nome} è in {s.classe}ª: la musica da camera è solo per 3ª, 4ª e 5ª.")
                continue
            risolti.append(s.id)
            membri[s.id].append(g.numero)
        g.studenti = risolti
        if not (MIN_GRUPPO_LMC <= len(g.studenti_raw) <= MAX_GRUPPO_LMC):
            err(dove, f"Il gruppo ha {len(g.studenti_raw)} studenti: devono essere da {MIN_GRUPPO_LMC} a {MAX_GRUPPO_LMC}.")
    for s in dati.studenti:
        if s.ore[2] == 0:
            continue
        gr = membri.get(s.id, [])
        dove = f"{FOGLIO_STUDENTI}, riga {s.riga} ({s.cognome} {s.nome}, {s.classe}ª)"
        if not gr:
            err(dove, f"Non è in nessun gruppo di musica da camera (foglio {FOGLIO_GRUPPI}).")
        elif len(gr) > 1:
            err(dove, f"È in più gruppi di musica da camera ({', '.join(map(str, gr))}): deve stare in uno solo.")

    # ── Docenti: ore richieste vs disponibilità ──
    # a chi non bastano le fasce dichiarate se ne aprono quante ne servono, segnalandolo
    problemi.extend(apri_fasce_mancanti(dati))
    ore_richieste = _ore_richieste(dati)
    for d in dati.docenti:
        dove = f"{FOGLIO_DOCENTI}: {d.nome}"
        n_x, n_a = len(d.fasce_x), len(d.fasce_a)
        richieste = ore_richieste.get(d.nome, 0)
        if richieste == 0 and d.ore_accomp == 0:
            avv(dove, "Nessuna lezione assegnata a questo docente.", "docenti senza nessuna lezione")
            continue
        if n_a > d.ore_accomp:
            err(dove, f"Ci sono {n_a} caselle con A ma le ore di accompagnamento indicate sono {d.ore_accomp}.")
        totale = richieste + d.ore_accomp
        if totale > N_FASCE:
            err(dove, f"Servono {richieste} ore di lezione + {d.ore_accomp} di accompagnamento = {totale}, "
                      f"ma le fasce pomeridiane della settimana sono {N_FASCE}. Va ridotto il carico o aggiunto un docente.")
        elif totale > n_x + n_a:
            err(dove, f"Servono {totale} ore ({richieste} lezione + {d.ore_accomp} accompagnamento) "
                      f"ma il docente ha solo {n_x + n_a} fasce disponibili.")
        elif totale == n_x + n_a:
            avv(dove, f"Le ore richieste ({totale}) coincidono esattamente con le fasce disponibili: nessun margine.",
                "docenti senza nessuna ora di margine")

    # ── Studente: fattibilità elementare con i suoi docenti ──
    for s in dati.studenti:
        dove = f"{FOGLIO_STUDENTI}, riga {s.riga} ({s.cognome} {s.nome}, {s.classe}ª)"
        giorni_ok = [g for g in range(5) if g not in s.giorni_non_disp]
        libere_stud = {f for f in range(N_FASCE) if s.libero(f)}
        for doc_nome, ore in ((s.doc1, s.ore[0]), (s.doc2, s.ore[1])):
            if not ore or doc_nome not in docenti:
                continue
            d = docenti[doc_nome]
            fasce = [f for f in d.fasce_x if f in libere_stud]
            if len(fasce) < ore:
                err(dove, f"Il docente {doc_nome} non ha abbastanza fasce disponibili nelle ore in cui lo studente può venire "
                          f"({len(fasce)} fasce per {ore} ore).")
            if ore == 2 and s.ore_consecutive:
                coppie = [f for f in fasce if f % N_ORE < N_ORE - 1 and (f + 1) in fasce]
                if not coppie:
                    err(dove, f"Ore consecutive = SI ma il docente {doc_nome} non ha due fasce consecutive libere nello stesso giorno.")
        if s.giorno_unico:
            # esiste un giorno in cui tutti i suoi docenti hanno disponibilità?
            docs = [docenti[n] for n in (s.doc1, s.doc2) if n in docenti]
            g_lmc = dati.gruppo_di(s.id)
            if g_lmc and g_lmc.docente in docenti:
                docs.append(docenti[g_lmc.docente])
            giorni_comuni = [g for g in giorni_ok if all(any(f // N_ORE == g for f in d.fasce_x) for d in docs)]
            if docs and not giorni_comuni:
                err(dove, "Giorno unico = SI ma non c'è nessun giorno in cui tutti i suoi docenti sono disponibili.")

    # ── Trasporti ──
    con_indirizzo = [s for s in dati.studenti if s.indirizzo_completo]
    con_trasporto = [s for s in con_indirizzo if s.trasporto is not None]
    if con_indirizzo and not con_trasporto:
        avv(FOGLIO_STUDENTI, f"{len(con_indirizzo)} studenti hanno l'indirizzo ma i tempi dei mezzi non sono stati calcolati: "
                             "usati i KM. Menu Orario → Aggiorna trasporti (serve internet).",
            "tempi dei mezzi non calcolati")
    elif con_indirizzo and len(con_trasporto) < len(con_indirizzo):
        mancanti = [s for s in con_indirizzo if s.trasporto is None]
        avv(FOGLIO_STUDENTI, f"Tempi dei mezzi mancanti o non aggiornati per {len(mancanti)} studenti "
                             f"({', '.join(s.cognome.title() for s in mancanti[:6])}{' …' if len(mancanti) > 6 else ''}): usati i KM.",
            "tempi dei mezzi non calcolati")

    # ── LMI ──
    for l in dati.lmi:
        if l.docente and l.docente not in docenti:
            avv(f"{FOGLIO_LMI}: {l.nome}", f"Il docente '{l.docente}' non è nel foglio {FOGLIO_DOCENTI} (solo per la stampa).")

    errori = [p for p in problemi if p.livello == "errore"]
    if errori:
        raise ProblemiError(problemi)
    return problemi


def riepilogo_fasce(f: int) -> str:
    return nome_fascia(f)


__all__ = ["controlla", "risolvi_studente", "ORE_PER_CLASSE"]
