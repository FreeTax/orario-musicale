#!/usr/bin/env python3
"""Test end-to-end approfonditi del comportamento non-GUI.

Esecuzione:  .venv/bin/python test_e2e.py
Copre: template/lettura, controlli, motore, export, trasporti (senza rete, con _get simulato).
Lavora sempre in cartelle temporanee; il file esempio/orario_2026-27.xlsx viene solo letto.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
import urllib.error
import re
import urllib.parse
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

from orario import controlli, lettura, motore, template, trasporti
from orario.costanti import (
    FOGLI_DATI, FOGLIO_ABBINAMENTI, FOGLIO_IMPEGNI,
    COLONNE_AULE,
    FASCE, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_LMI, FOGLIO_PARAMETRI, FOGLIO_STUDENTI,
    FOGLIO_TRASPORTI, GIORNI, MINUTI_SENZA_MEZZO, N_GIORNI, N_ORE, ORE, TIPO_ACCOMP, TIPO_LMC,
    TIPO_LMI, TIPO_STRUM1, TIPO_STRUM2, fascia, giorno_ora,
)
from orario.export import esporta_tutto
from orario.lettura import Tabella
from orario.modello import DatiInput, GruppoLMC, Parametri, ProblemiError, Trasporto
from orario.template import (
    COLONNE_ABBINAMENTI, COLONNE_DOCENTI, COLONNE_GRUPPI, COLONNE_IMPEGNI, COLONNE_LMI,
    COLONNE_PARAMETRI, COLONNE_STUDENTI,
    PARAMETRI_DEFAULT,
)
from orario.trasporti import Coordinate, TrasportiError

RADICE = Path(__file__).resolve().parent
ESEMPIO = RADICE / "esempio" / "orario_2026-27.xlsx"


# ── Helper per costruire dati a mano ─────────────────────────────────────────

def _riga(colonne: list[str], valori: dict) -> list[str]:
    return [str(valori.get(c, "")) for c in colonne]


def stud(cognome, classe=1, nome="MARIO", km=None, s1="PIANOFORTE", d1="Bianchi",
         s2="VIOLINO", d2="Verdi", cons="NO", unico="NO", giorni_no="", note="",
         comune="", indirizzo="", civico="") -> dict:
    return {
        "Classe": classe, "Cognome": cognome, "Nome": nome, "Comune": comune,
        "Indirizzo": indirizzo, "Civico": civico, "KM": "" if km is None else km,
        "Strumento 1": s1, "Docente 1": d1, "Strumento 2": s2, "Docente 2": d2,
        "Ore consecutive (SI/NO)": cons, "Giorno unico (SI/NO)": unico,
        "Giorni NON disponibili": giorni_no, "Note": note,
    }


def doc(nome, fasce=(), accomp=0, aula="M1", strumenti="PIANOFORTE", note="") -> dict:
    """`fasce`: iterabile di indici 0..19 (X) oppure dict indice -> 'X'/'A'."""
    d = {"Docente": nome, "Strumento/i": strumenti,
         "Ore accompagnamento": accomp, "Note": note}
    for nome_colonna in COLONNE_AULE:      # una aula per giorno
        d[nome_colonna] = aula
    mappa = fasce if isinstance(fasce, dict) else {f: "X" for f in fasce}
    for f, v in mappa.items():
        d[FASCE[f]] = v
    return d


def gruppo(numero, docente, *studenti, note="") -> dict:
    d = {"Gruppo": numero, "Docente": docente, "Note": note}
    for i, s in enumerate(studenti, start=1):
        d[f"Studente {i}"] = s
    return d


def tabelle(studenti=(), docenti=(), gruppi=(), lmi=(), parametri=None, impegni=(),
            abbinamenti=()) -> dict[str, Tabella]:
    par = {n: v for n, v, _ in PARAMETRI_DEFAULT}
    par.update(parametri or {})
    return {
        FOGLIO_STUDENTI: Tabella(FOGLIO_STUDENTI, list(COLONNE_STUDENTI),
                                 [_riga(COLONNE_STUDENTI, s) for s in studenti]),
        FOGLIO_DOCENTI: Tabella(FOGLIO_DOCENTI, list(COLONNE_DOCENTI),
                                [_riga(COLONNE_DOCENTI, d) for d in docenti]),
        FOGLIO_GRUPPI: Tabella(FOGLIO_GRUPPI, list(COLONNE_GRUPPI),
                               [_riga(COLONNE_GRUPPI, g) for g in gruppi]),
        FOGLIO_IMPEGNI: Tabella(FOGLIO_IMPEGNI, list(COLONNE_IMPEGNI),
                                [_riga(COLONNE_IMPEGNI, i) for i in impegni]),
        FOGLIO_ABBINAMENTI: Tabella(FOGLIO_ABBINAMENTI, list(COLONNE_ABBINAMENTI),
                                    [_riga(COLONNE_ABBINAMENTI, a) for a in abbinamenti]),
        FOGLIO_LMI: Tabella(FOGLIO_LMI, list(COLONNE_LMI),
                            [_riga(COLONNE_LMI, l) for l in lmi]),
        FOGLIO_PARAMETRI: Tabella(FOGLIO_PARAMETRI, list(COLONNE_PARAMETRI),
                                  [[n, str(v), ""] for n, v in par.items()]),
    }


def dati_di(studenti=(), docenti=(), gruppi=(), lmi=(), parametri=None,
            percorso=None, timeout=10, impegni=(), abbinamenti=()) -> DatiInput:
    par = {"Tempo massimo di calcolo (secondi)": timeout}
    par.update(parametri or {})
    d = lettura.costruisci_dati(tabelle(studenti, docenti, gruppi, lmi, par, impegni, abbinamenti), percorso)
    return d


def abbinamento(docente, studente, giorno="Lunedì", ora="13:30", tipo="", note="") -> dict:
    """Riga del foglio Abbinamenti fissi."""
    return {"Docente": docente, "Studente": studente, "Tipo di lezione": tipo,
            "Giorno": giorno, "Ora": ora, "Note": note}


def lab(nome, docente, *studenti, classi="1ª", aula="M9", quando="Martedì 4ª-5ª ora", note="") -> dict:
    """Riga del foglio LMI."""
    return {"Laboratorio": nome, "Classi": classi, "Docente": docente, "Aula": aula,
            "Giorno e ora (mattino)": quando,
            "Studenti (Cognome Nome, separati da virgola)": ", ".join(studenti), "Note": note}


def impegno(cognome, nome="MARIO", classe=1, fasce=(), note="") -> dict:
    """Riga del foglio Impegni: le fasce indicate sono ore in cui il ragazzo NON c'è."""
    d = {"Classe": classe, "Cognome": cognome, "Nome": nome, "Note": note}
    for f in fasce:
        d[FASCE[f]] = "X"
    return d


def tutte(giorni=range(5), ore=range(4)) -> list[int]:
    return [fascia(g, o) for g in giorni for o in ore]


class Caso(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="orario_test_"))
        self.addCleanup(shutil.rmtree, self.dir, True)

    def nuovo_file(self, nome="input.xlsx") -> Path:
        return template.crea_nuovo_file(self.dir / nome)


# ═════════════════════════════════════════════════════════════════════════════
# 1. Template e file di input
# ═════════════════════════════════════════════════════════════════════════════

class TestTemplate(Caso):

    def test_sei_fogli_e_colonne(self):
        p = self.nuovo_file()
        wb = load_workbook(p)
        self.assertEqual(wb.sheetnames,
                         ["Istruzioni"] + list(FOGLI_DATI))
        atteso = {FOGLIO_STUDENTI: COLONNE_STUDENTI, FOGLIO_DOCENTI: COLONNE_DOCENTI,
                  FOGLIO_GRUPPI: COLONNE_GRUPPI, FOGLIO_LMI: COLONNE_LMI,
                  FOGLIO_PARAMETRI: COLONNE_PARAMETRI}
        for foglio, colonne in atteso.items():
            got = [c.value for c in wb[foglio][1]][:len(colonne)]
            self.assertEqual(got, list(colonne), foglio)
        # le 20 colonne di disponibilità ci sono tutte, di seguito e nell'ordine di FASCE
        intest = [c.value for c in wb[FOGLIO_DOCENTI][1]]
        inizio = intest.index(FASCE[0])
        self.assertEqual(intest[inizio:inizio + len(FASCE)], FASCE)
        # e prima ci sono le cinque aule, una per giorno
        self.assertEqual([c for c in intest if str(c).startswith("Aula ")], list(COLONNE_AULE))

    def test_righe_esempio_fanno_fallire(self):
        p = self.nuovo_file()
        with self.assertRaises(ProblemiError) as cm:
            lettura.leggi_dati(p)
        messaggi = [str(x) for x in cm.exception.problemi]
        testo = "\n".join(messaggi)
        # cita foglio e riga per studenti, docenti e gruppi
        for foglio, righe in ((FOGLIO_STUDENTI, (2, 3)), (FOGLIO_DOCENTI, (2, 3)),
                              (FOGLIO_GRUPPI, (2,))):
            for r in righe:
                self.assertTrue(any(f"{foglio}, riga {r}" in m for m in messaggi),
                                f"manca '{foglio}, riga {r}' in:\n{testo}")
        self.assertIn("esempio", testo.lower())

    def test_round_trip_tabelle(self):
        p = self.nuovo_file()
        t = lettura.leggi_tabelle(p)
        ts = t[FOGLIO_STUDENTI]
        ts.righe = [
            _riga(COLONNE_STUDENTI, stud("DEGL’INNOCENTI", nome="NICCOLÒ", km=27, comune="Pistoia",
                                         indirizzo="Via Nazario Sauro", civico="283",
                                         note="perché è così: “virgolette”")),
            _riga(COLONNE_STUDENTI, stud("ROSSI", classe=3, nome="", km="12.5", d1="Bianchi",
                                         d2="", giorni_no="Mar, Gio")),
        ]
        lettura.salva_tabelle(p, t)
        t2 = lettura.leggi_tabelle(p)
        self.assertEqual(t2[FOGLIO_STUDENTI].intestazione, ts.intestazione)
        self.assertEqual(t2[FOGLIO_STUDENTI].righe, ts.righe)
        # idempotenza: un secondo giro non cambia nulla
        lettura.salva_tabelle(p, t2)
        self.assertEqual(lettura.leggi_tabelle(p)[FOGLIO_STUDENTI].righe, ts.righe)
        # celle vuote restano vuote (non diventano "None")
        self.assertEqual(t2[FOGLIO_STUDENTI].righe[1][COLONNE_STUDENTI.index("Nome")], "")

    def test_decimale_con_virgola_diventa_punto(self):
        """Comportamento noto: '12,5' viene salvato come numero e riletto '12.5' (stabile dopo)."""
        p = self.nuovo_file()
        t = lettura.leggi_tabelle(p)
        t[FOGLIO_STUDENTI].righe = [_riga(COLONNE_STUDENTI, stud("BIANCHI", km="12,5"))]
        lettura.salva_tabelle(p, t)
        t2 = lettura.leggi_tabelle(p)
        letto = t2[FOGLIO_STUDENTI]
        self.assertEqual(letto.valore(letto.righe[0], "KM"), "12.5")
        # ma è comunque interpretato come 12.5 dal parser
        t2[FOGLIO_DOCENTI] = tabelle(docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])[FOGLIO_DOCENTI]
        t2[FOGLIO_GRUPPI].righe = []
        t2[FOGLIO_LMI].righe = []
        d = lettura.costruisci_dati(t2)
        self.assertEqual(d.studenti[0].km, 12.5)

    def test_salva_tabelle_conserva_tendine_e_intestazione(self):
        p = self.nuovo_file()
        prima = {f: len(load_workbook(p)[f].data_validations.dataValidation)
                 for f in (FOGLIO_STUDENTI, FOGLIO_DOCENTI, FOGLIO_GRUPPI)}
        self.assertTrue(all(v > 0 for v in prima.values()), prima)
        t = lettura.leggi_tabelle(p)
        t[FOGLIO_STUDENTI].righe = [_riga(COLONNE_STUDENTI, stud("VERDI"))]
        lettura.salva_tabelle(p, t)
        wb = load_workbook(p)
        dopo = {f: len(wb[f].data_validations.dataValidation) for f in prima}
        self.assertEqual(prima, dopo)
        self.assertEqual([c.value for c in wb[FOGLIO_STUDENTI][1]][:len(COLONNE_STUDENTI)],
                         list(COLONNE_STUDENTI))

    def test_aggiorna_struttura(self):
        t = tabelle(studenti=[stud("ROSSI", km=10, note="nota")],
                    docenti=[doc("Bianchi", tutte())])
        ts = t[FOGLIO_STUDENTI]
        # togli Comune / Indirizzo / Civico come nei file vecchi
        for colonna in ("Comune", "Indirizzo", "Civico"):
            i = ts.intestazione.index(colonna)
            ts.intestazione.pop(i)
            for r in ts.righe:
                r.pop(i)
        tp = t[FOGLIO_PARAMETRI]
        tp.righe = [r for r in tp.righe
                    if "indirizzo della scuola" not in r[0].lower()
                    and "soglia minuti" not in r[0].lower()]
        modifiche = lettura.aggiorna_struttura(t)
        self.assertEqual(ts.intestazione, list(COLONNE_STUDENTI))
        self.assertEqual(len(ts.righe[0]), len(COLONNE_STUDENTI))
        # i dati sono ancora allineati alle colonne giuste
        self.assertEqual(ts.valore(ts.righe[0], "Cognome"), "ROSSI")
        self.assertEqual(ts.valore(ts.righe[0], "KM"), "10")
        self.assertEqual(ts.valore(ts.righe[0], "Note"), "nota")
        self.assertEqual(ts.valore(ts.righe[0], "Comune"), "")
        chiavi = [r[0].lower() for r in tp.righe]
        self.assertIn("indirizzo della scuola", chiavi)
        self.assertTrue(any("soglia minuti" in k for k in chiavi))
        self.assertEqual(len(modifiche), 5, modifiche)
        # idempotente
        self.assertEqual(lettura.aggiorna_struttura(t), [])

    # ── valori sporchi ──
    def _errori(self, **kw) -> str:
        with self.assertRaises(ProblemiError) as cm:
            dati_di(**kw)
        return "\n".join(str(p) for p in cm.exception.problemi)

    def test_km_con_virgola_ok(self):
        d = dati_di(studenti=[stud("ROSSI", km="27,5")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        self.assertEqual(d.studenti[0].km, 27.5)

    def test_km_non_numerico(self):
        t = self._errori(studenti=[stud("ROSSI", km="circa 12")], docenti=[doc("Bianchi", tutte())])
        self.assertIn("KM 'circa 12' non è un numero", t)
        self.assertIn("riga 2", t)

    def test_classe_fuori_intervallo(self):
        t = self._errori(studenti=[stud("ROSSI", classe=6)], docenti=[doc("Bianchi", tutte())])
        self.assertIn("Classe '6' non valida", t)

    def test_giorno_non_riconosciuto(self):
        t = self._errori(studenti=[stud("ROSSI", giorni_no="Sabato, Lun")],
                         docenti=[doc("Bianchi", tutte())])
        self.assertIn("Giorni non riconosciuti", t)
        self.assertIn("Sabato", t)
        self.assertIn("Usare Lun, Mar, Mer, Gio, Ven", t)

    def test_x_minuscola_accettata(self):
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[{**doc("Bianchi"), FASCE[0]: "x", FASCE[1]: "a", "Ore accompagnamento": 1},
                             doc("Verdi", tutte())])
        b = d.docente("Bianchi")
        self.assertEqual(b.disponibilita, {0: "X", 1: "A"})

    def test_valore_disponibilita_non_ammesso(self):
        t = self._errori(studenti=[stud("ROSSI")],
                         docenti=[{**doc("Bianchi"), FASCE[3]: "SI"}])
        self.assertIn("Lun 16:30", t)
        self.assertIn("ammessi solo X, A o vuoto", t)

    def test_ore_accomp_non_intera(self):
        t = self._errori(docenti=[{**doc("Bianchi", tutte()), "Ore accompagnamento": "due"}])
        self.assertIn("non è un numero intero", t)

    def test_classe5_ignora_secondo_strumento(self):
        d = dati_di(studenti=[stud("ROSSI", classe=5, s2="VIOLINO", d2="Verdi")],
                    docenti=[doc("Bianchi", tutte())])
        self.assertEqual((d.studenti[0].strum2, d.studenti[0].doc2), ("", ""))
        self.assertEqual(d.studenti[0].ore, (2, 0, 1))


# ═════════════════════════════════════════════════════════════════════════════
# 2. Controlli
# ═════════════════════════════════════════════════════════════════════════════

class TestControlli(Caso):

    def errori(self, dati) -> str:
        with self.assertRaises(ProblemiError) as cm:
            controlli.controlla(dati)
        return "\n".join(str(p) for p in cm.exception.errori)

    def test_docente_inesistente(self):
        d = dati_di(studenti=[stud("ROSSI", d1="Ignoto", d2="Verdi")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        t = self.errori(d)
        self.assertIn("Il docente 'Ignoto' del 1° strumento non è nel foglio Docenti", t)
        self.assertIn("Studenti, riga 2 (ROSSI MARIO, 1ª)", t)

    def test_docente_mancante(self):
        d = dati_di(studenti=[stud("ROSSI", d1="", d2="")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        t = self.errori(d)
        self.assertIn("Manca il docente del 1° strumento", t)
        self.assertIn("Manca il docente del 2° strumento", t)

    def test_studente_duplicato(self):
        d = dati_di(studenti=[stud("ROSSI"), stud("ROSSI")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        self.assertIn("Studente presente due volte", self.errori(d))

    def test_docente_ripetuto(self):
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[doc("Bianchi", tutte()), doc("Bianchi", tutte()), doc("Verdi", tutte())])
        t = self.errori(d)
        self.assertIn("Docenti: Bianchi", t)
        self.assertIn("Docente ripetuto due volte", t)

    def test_gruppo_studente_inesistente(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3), stud("NERI", classe=3)],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Bianchi", "Rossi", "Manzi")])
        t = self.errori(d)
        self.assertIn("Studente 'Manzi'", t)
        self.assertIn("non trovato nel foglio Studenti", t)
        self.assertIn("Gruppi LMC, gruppo 1", t)

    def test_gruppo_con_studente_del_biennio(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3), stud("NERI", classe=2, nome="ANNA")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Bianchi", "Rossi", "Neri")])
        t = self.errori(d)
        self.assertIn("è in 2ª: la musica da camera è solo per 3ª, 4ª e 5ª", t)

    def test_gruppo_troppo_piccolo(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3)],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Bianchi", "Rossi")])
        t = self.errori(d)
        self.assertIn("Il gruppo ha 1 studenti: devono essere da 2 a 5", t)

    def test_gruppo_troppo_grande(self):
        cognomi = ["A" * 3, "BEE", "CIT", "DOR", "ELF", "FAB"]
        d = dati_di(studenti=[stud(c, classe=3) for c in cognomi],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Bianchi", *cognomi[:5])])
        # la sesta colonna non esiste nel foglio: si costruisce il gruppo a mano
        d.gruppi[0].studenti_raw = list(cognomi)
        t = self.errori(d)
        self.assertIn("Il gruppo ha 6 studenti: devono essere da 2 a 5", t)

    def test_studente_triennio_senza_gruppo(self):
        d = dati_di(studenti=[stud("ROSSI", classe=4)],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        self.assertIn("Non è in nessun gruppo di musica da camera", self.errori(d))

    def test_studente_in_due_gruppi(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3), stud("NERI", classe=3, nome="ANNA"),
                              stud("GIALLI", classe=3, nome="LUCA")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Bianchi", "Rossi", "Neri"),
                            gruppo(2, "Bianchi", "Rossi", "Gialli")])
        t = self.errori(d)
        self.assertIn("È in più gruppi di musica da camera (1, 2)", t)

    def test_ore_richieste_oltre_disponibilita(self):
        """Fasce insufficienti: il programma ne apre quante ne servono e lo segnala."""
        d = dati_di(studenti=[stud(c, classe=3) for c in ("ROSSI", "NERI", "GIALLI")],
                    docenti=[doc("Bianchi", [0]), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Verdi", "Rossi", "Neri", "Gialli")])
        avvisi = controlli.controlla(d)
        t = "\n".join(str(a) for a in avvisi)
        self.assertIn("Docenti: Bianchi", t)
        self.assertIn("Servivano 3 ore ma erano dichiarate disponibili solo 1", t)
        self.assertIn("Da concordare con il docente", t)
        self.assertEqual(len(d.docente("Bianchi").disponibilita), 3, "le fasce aperte sono quelle mancanti")
        self.assertIn(0, d.docente("Bianchi").disponibilita, "la fascia già dichiarata resta")

    def test_docente_senza_x(self):
        """Nessuna disponibilità: si aprono tutte le ore che servono, con la segnalazione."""
        d = dati_di(studenti=[stud("ROSSI", classe=3), stud("NERI", classe=3, nome="ANNA", d1="Verdi", d2="Verdi")],
                    docenti=[doc("Bianchi", []), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Verdi", "Rossi", "Neri")])
        avvisi = controlli.controlla(d)
        t = "\n".join(str(a) for a in avvisi)
        self.assertIn("Docenti: Bianchi", t)
        self.assertIn("erano dichiarate disponibili solo 0 fasce", t)
        self.assertEqual(len(d.docente("Bianchi").disponibilita), 1)

    def test_studente_senza_giorni(self):
        d = dati_di(studenti=[stud("ROSSI", giorni_no="Lun, Mar, Mer, Gio, Ven")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        self.assertIn("Lo studente non è disponibile in nessun giorno", self.errori(d))

    def test_giorno_unico_impossibile(self):
        d = dati_di(studenti=[stud("ROSSI", unico="SI", d1="Bianchi", d2="Verdi")],
                    docenti=[doc("Bianchi", tutte(giorni=[0])), doc("Verdi", tutte(giorni=[1]))])
        t = self.errori(d)
        self.assertIn("Giorno unico = SI ma non c'è nessun giorno in cui tutti i suoi docenti sono disponibili", t)

    def test_ore_consecutive_impossibili(self):
        # Bianchi ha X solo alla 4ª fascia di ogni giorno: nessuna coppia adiacente
        d = dati_di(studenti=[stud("ROSSI", cons="SI")],
                    docenti=[doc("Bianchi", [fascia(g, 3) for g in range(5)]), doc("Verdi", tutte())])
        t = self.errori(d)
        self.assertIn("non ha due fasce consecutive libere nello stesso giorno", t)

    def test_avvisi_non_bloccano(self):
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte()), doc("Ozio", tutte())])
        avvisi = controlli.controlla(d)
        testo = "\n".join(str(a) for a in avvisi)
        self.assertTrue(all(a.livello == "avviso" for a in avvisi))
        self.assertIn("Nessuna lezione assegnata a questo docente", testo)
        self.assertIn("Né KM né tempi dei mezzi", testo)

    def test_avviso_ore_consecutive_ignorate(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3, cons="SI")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    gruppi=[gruppo(1, "Verdi", "Rossi", "Neri")])
        d.gruppi[0].studenti_raw = ["Rossi"]
        d.gruppi.append(GruppoLMC(numero=2, docente="Verdi", studenti=[], studenti_raw=["Rossi"], riga=3))
        with self.assertRaises(ProblemiError):
            controlli.controlla(d)  # in due gruppi: errore atteso
        d2 = dati_di(studenti=[stud("ROSSI", classe=3, cons="SI"), stud("NERI", classe=3, nome="ANNA")],
                     docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                     gruppi=[gruppo(1, "Verdi", "Rossi", "Neri")])
        testo = "\n".join(str(a) for a in controlli.controlla(d2))
        self.assertIn("Ore consecutive = SI ma la classe ha una sola ora di 1° strumento", testo)

    def test_avviso_a_oltre_ore_accomp(self):
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[{**doc("Bianchi", tutte()), FASCE[3]: "A", "Ore accompagnamento": 0},
                             doc("Verdi", tutte())])
        self.assertIn("caselle con A ma le ore di accompagnamento indicate sono 0", self.errori(d))

    # ── risolvi_studente ──
    def test_risolvi_studente(self):
        d = dati_di(studenti=[stud("ROSSI", nome="MARIO"), stud("ROSSI", nome="ANNA"),
                              stud("DEGL’INNOCENTI", nome="NICCOLÒ"),
                              stud("DE LUCA", nome="GIORGIA")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())])
        S = d.studenti
        r = controlli.risolvi_studente
        self.assertIs(r("Degl'Innocenti", S)[0], S[2])          # apostrofo diritto → tipografico
        self.assertIs(r("DEGL’INNOCENTI", S)[0], S[2])
        self.assertIs(r("De Luca", S)[0], S[3])                 # cognome composto
        self.assertIs(r("Rossi Anna", S)[0], S[1])              # cognome + nome
        self.assertIs(r("Rossi M.", S)[0], S[0])                # cognome + iniziale
        s, motivo = r("Rossi", S)                                # omonimi
        self.assertIsNone(s)
        self.assertIn("scrivere anche il nome", motivo)
        self.assertIn("ROSSI MARIO", motivo)
        self.assertIn("(1ª)", motivo)
        self.assertIsNone(r("Manzi", S)[0])
        self.assertIn("non trovato", r("Manzi", S)[1])
        self.assertIn("vuoto", r("   ", S)[1])


# ═════════════════════════════════════════════════════════════════════════════
# 3. Motore
# ═════════════════════════════════════════════════════════════════════════════

def controlla_invarianti(caso: unittest.TestCase, orario) -> None:
    """Tutte le regole rigide di sez. 6 verificate sull'orario prodotto."""
    dati = orario.dati
    docenti = {d.nome: d for d in dati.docenti}
    # ore per studente
    for s in dati.studenti:
        mie = orario.lezioni_di(s.id)
        h1, h2, hl = s.ore
        conta = {t: sum(1 for l in mie if l.tipo == t) for t in (TIPO_STRUM1, TIPO_STRUM2, TIPO_LMC)}
        caso.assertEqual((conta[TIPO_STRUM1], conta[TIPO_STRUM2], conta[TIPO_LMC]), (h1, h2, hl),
                         f"ore sbagliate per {s.id}")
        # mai due lezioni nella stessa fascia
        fasce = [l.fascia for l in mie]
        caso.assertEqual(len(fasce), len(set(fasce)), f"{s.id} sovrapposto")
        # giorni non disponibili
        for l in mie:
            caso.assertNotIn(l.giorno, s.giorni_non_disp, f"{s.id} il {GIORNI[l.giorno]}")
        if s.giorno_unico:
            caso.assertEqual(len({l.giorno for l in mie}), 1, f"{s.id} giorno unico")
        if s.ore_consecutive and h1 == 2:
            a, b = sorted((l for l in mie if l.tipo == TIPO_STRUM1), key=lambda l: l.fascia)
            caso.assertEqual((a.giorno, b.ora), (b.giorno, a.ora + 1))
            caso.assertTrue(b.due_ore and not a.due_ore)
        caso.assertLessEqual(orario.rientri_di(s.id), dati.parametri.max_rientri_vicini)
    # docenti: una lezione per fascia, solo fasce con X (o A per l'accompagnamento)
    visti = set()
    for l in orario.lezioni:
        caso.assertNotIn((l.docente, l.fascia), visti, f"{l.docente} doppio {l.fascia}")
        visti.add((l.docente, l.fascia))
        v = docenti[l.docente].disponibilita.get(l.fascia)
        if l.tipo == TIPO_ACCOMP:
            caso.assertIn(v, ("X", "A"))
        else:
            caso.assertEqual(v, "X", f"{l.docente} {l.fascia} disponibilità {v}")
    # accompagnamento: numero esatto, A coperte, sempre in coda alle lezioni del giorno
    for d in dati.docenti:
        acc = [l for l in orario.lezioni if l.docente == d.nome and l.tipo == TIPO_ACCOMP]
        lez = [l for l in orario.lezioni if l.docente == d.nome and l.tipo != TIPO_ACCOMP]
        caso.assertEqual(len(acc), d.ore_accomp, f"{d.nome} accomp")
        for f in d.fasce_a:
            caso.assertEqual(sum(1 for l in acc if l.fascia == f), 1,
                             f"la A di {d.nome} in {f} non è coperta")
        for g in range(N_GIORNI):
            oa = [l.ora for l in acc if l.giorno == g]
            ol = [l.ora for l in lez if l.giorno == g]
            if oa and ol:
                caso.assertGreater(min(oa), max(ol), f"{d.nome}: accomp prima di una lezione il {GIORNI[g]}")
    # gruppi LMC: una sola lezione, stessa fascia, docente giusto
    for g in dati.gruppi:
        lg = [l for l in orario.lezioni if l.tipo == TIPO_LMC and l.gruppo == g.numero]
        caso.assertEqual(len(lg), 1)
        caso.assertEqual(lg[0].docente, g.docente)
        caso.assertEqual(set(lg[0].studenti), set(g.studenti))


class TestMotore(Caso):

    def test_caso_minimo(self):
        d = dati_di(studenti=[stud("ROSSI"), stud("NERI", nome="ANNA")],
                    docenti=[doc("Bianchi", tutte(giorni=[0, 1])),
                             doc("Verdi", tutte(giorni=[0, 1]), strumenti="VIOLINO")],
                    timeout=10)
        o = motore.calcola(d)
        self.assertEqual(len(o.lezioni), 6)  # 2 studenti × (2 + 1) ore
        controlla_invarianti(self, o)

    def test_accompagnamento_in_coda(self):
        d = dati_di(studenti=[stud("ROSSI"), stud("NERI", nome="ANNA")],
                    docenti=[doc("Bianchi", tutte(giorni=[0, 1]), accomp=2),
                             doc("Verdi", tutte(giorni=[0, 1]))],
                    timeout=15)
        o = motore.calcola(d)
        acc = [l for l in o.lezioni if l.tipo == TIPO_ACCOMP]
        self.assertEqual(len(acc), 2)
        controlla_invarianti(self, o)

    def test_fascia_A_contiene_un_accomp(self):
        f_a = fascia(0, 3)
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[{**doc("Bianchi", {f: "X" for f in tutte(giorni=[0, 1])}, accomp=1),
                              FASCE[f_a]: "A"},
                             doc("Verdi", tutte(giorni=[0, 1]))],
                    timeout=15)
        o = motore.calcola(d)
        acc = [l for l in o.lezioni if l.tipo == TIPO_ACCOMP]
        self.assertEqual([l.fascia for l in acc], [f_a])
        controlla_invarianti(self, o)

    def test_gruppo_lmc_stessa_fascia(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3), stud("NERI", classe=4, nome="ANNA"),
                              stud("GIALLI", classe=5, nome="LUCA")],
                    docenti=[doc("Bianchi", tutte(giorni=[0, 1, 2])),
                             doc("Verdi", tutte(giorni=[0, 1, 2])),
                             doc("Camera", tutte(giorni=[0, 1, 2]))],
                    gruppi=[gruppo(1, "Camera", "Rossi", "Neri", "Gialli")],
                    timeout=15)
        controlli.controlla(d)
        o = motore.calcola(d)
        lmc = [l for l in o.lezioni if l.tipo == TIPO_LMC]
        self.assertEqual(len(lmc), 1)
        self.assertEqual(len(lmc[0].studenti), 3)
        for sid in lmc[0].studenti:
            self.assertIn(lmc[0], o.lezioni_di(sid))
        controlla_invarianti(self, o)

    def test_giorni_non_disponibili_e_giorno_unico(self):
        d = dati_di(studenti=[stud("ROSSI", giorni_no="Lun, Mar"), stud("NERI", nome="ANNA", unico="SI")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    timeout=15)
        o = motore.calcola(d)
        rossi = o.lezioni_di("ROSSI MARIO")
        self.assertTrue(all(l.giorno >= 2 for l in rossi))
        self.assertEqual(len({l.giorno for l in o.lezioni_di("NERI ANNA")}), 1)
        controlla_invarianti(self, o)

    def test_ore_consecutive(self):
        d = dati_di(studenti=[stud("ROSSI", cons="SI")],
                    docenti=[doc("Bianchi", tutte(giorni=[0])), doc("Verdi", tutte(giorni=[0]))],
                    timeout=10)
        o = motore.calcola(d)
        s1 = sorted((l for l in o.lezioni if l.tipo == TIPO_STRUM1), key=lambda l: l.fascia)
        self.assertEqual(len(s1), 2)
        self.assertEqual(s1[0].giorno, s1[1].giorno)
        self.assertEqual(s1[1].ora, s1[0].ora + 1)
        self.assertFalse(s1[0].due_ore)
        self.assertTrue(s1[1].due_ore)
        controlla_invarianti(self, o)

    def test_impossibile_docente_senza_disponibilita(self):
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[doc("Bianchi", []), doc("Verdi", tutte())], timeout=10)
        with self.assertRaises(ProblemiError) as cm:
            motore.calcola(d)
        t = "\n".join(str(p) for p in cm.exception.problemi)
        self.assertIn("Rossi Mario", t)
        self.assertIn("Bianchi", t)
        self.assertIn("non ha nessuna X", t)

    def test_impossibile_unica_fascia_condivisa(self):
        d = dati_di(studenti=[stud("ROSSI", classe=3, d1="Stretto"),
                              stud("NERI", classe=3, nome="ANNA", d1="Stretto")],
                    docenti=[doc("Stretto", [fascia(0, 0)]),
                             doc("Verdi", tutte(giorni=[2, 3])),
                             doc("Camera", tutte(giorni=[2, 3]))],
                    gruppi=[gruppo(1, "Camera", "Rossi", "Neri")],
                    timeout=15)
        # qui interessa la diagnosi del motore: i controlli fermerebbero prima questo caso,
        # quindi risolvo i gruppi a mano come farebbe controlla()
        d.gruppi[0].studenti = ["ROSSI MARIO", "NERI ANNA"]
        with self.assertRaises(ProblemiError) as cm:
            motore.calcola(d)
        t = "\n".join(str(p) for p in cm.exception.problemi)
        self.assertIn("Stretto", t)
        self.assertTrue("Rossi Mario" in t or "Neri Anna" in t, t)
        self.assertIn("impossibile collocare", t.lower())

    def test_gruppi_non_risolti_senza_controlla(self):
        """Senza controlli.controlla() i gruppi LMC non sono risolti: il motore lo dice chiaramente.

        I cognomi scritti nei gruppi vengono tradotti in studenti da controlla(), che ha questo
        effetto collaterale. Il motore si rifiuta di calcolare invece di produrre gruppi vuoti.
        """
        d = dati_di(studenti=[stud("ROSSI", classe=3), stud("NERI", classe=3, nome="ANNA")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte()), doc("Camera", tutte())],
                    gruppi=[gruppo(1, "Camera", "Rossi", "Neri")], timeout=10)
        self.assertEqual(d.gruppi[0].studenti, [])          # ancora da risolvere
        self.assertEqual(d.gruppi[0].studenti_raw, ["Rossi", "Neri"])
        with self.assertRaises(ProblemiError) as cm:
            motore.calcola(d)
        t = "\n".join(str(p) for p in cm.exception.problemi)
        self.assertIn("non sono stati controllati", t)
        self.assertIn("controlla", t)
        controlli.controlla(d)
        self.assertEqual(len(d.gruppi[0].studenti), 2)
        controlla_invarianti(self, motore.calcola(d))

    def test_rientri_entro_il_massimo(self):
        d = dati_di(studenti=[stud("ROSSI", km=40)],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    parametri={"Max rientri per ragazzo": 2, "Max rientri per chi abita vicino": 2},
                    timeout=10)
        o = motore.calcola(d)
        self.assertLessEqual(o.rientri_di("ROSSI MARIO"), 2)

    def test_trasporti_lontano_prima(self):
        """Due studenti identici: chi ha 120 minuti di ritorno va in una fascia <= dell'altro."""
        d = dati_di(studenti=[stud("VICINO", classe=3, nome="V", d1="Stretto", d2="Largo"),
                              stud("LONTANO", classe=3, nome="L", d1="Stretto", d2="Largo")],
                    docenti=[doc("Stretto", [fascia(0, 0), fascia(0, 1)]),
                             doc("Largo", tutte(giorni=[1]))],
                    gruppi=[gruppo(1, "Largo", "Vicino", "Lontano")],
                    timeout=15)
        d.studente("VICINO V").trasporto = Trasporto([15] * 4, ["15:00"] * 4, ["bus"] * 4,
                                                     indirizzo_usato="", esito="OK")
        d.studente("LONTANO L").trasporto = Trasporto([120] * 4, ["17:00"] * 4, ["bus"] * 4,
                                                      indirizzo_usato="", esito="OK")
        controlli.controlla(d)
        o = motore.calcola(d)
        s1 = {l.studenti[0]: l for l in o.lezioni if l.tipo == TIPO_STRUM1}
        self.assertLessEqual(s1["LONTANO L"].ora, s1["VICINO V"].ora)
        controlla_invarianti(self, o)


class TestAbbinamentiFissi(Caso):
    """Il foglio delle lezioni già decise: ragazzi, laboratori LMI e ore di 1° strumento."""

    def test_lezione_bloccata_nella_sua_fascia(self):
        f = fascia(1, 2)  # Martedì 15:30
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[doc("Bianchi", tutte(giorni=[0, 1])), doc("Verdi", tutte(giorni=[0, 1]))],
                    abbinamenti=[abbinamento("Verdi", "Rossi", "Martedì", "15:30", "2° strumento")],
                    timeout=10)
        controlli.controlla(d)
        o = motore.calcola(d)
        s2 = [l for l in o.lezioni if l.tipo == TIPO_STRUM2]
        self.assertEqual([l.fascia for l in s2], [f])
        controlla_invarianti(self, o)

    def test_due_ore_di_primo_strumento_lontane_e_in_ordine_sparso(self):
        """Le 2 ore di 1° strumento non sono attaccate: si mettono dove dicono gli abbinamenti.

        Le due righe sono scritte apposta al contrario (prima il venerdì, poi il lunedì)."""
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
                    abbinamenti=[abbinamento("Bianchi", "Rossi", "Venerdì", "16:30", "1° strumento"),
                                 abbinamento("Bianchi", "Rossi", "Lunedì", "13:30", "1° strumento")],
                    timeout=15)
        controlli.controlla(d)
        o = motore.calcola(d)
        s1 = sorted(l.fascia for l in o.lezioni if l.tipo == TIPO_STRUM1)
        self.assertEqual(s1, [fascia(0, 0), fascia(4, 3)])
        controlla_invarianti(self, o)

    def test_primo_strumento_puo_stare_in_giorni_diversi(self):
        """Senza abbinamenti niente obbliga le 2 ore a essere attaccate: qui non potrebbero esserlo."""
        d = dati_di(studenti=[stud("ROSSI")],
                    docenti=[doc("Bianchi", [fascia(0, 0), fascia(1, 0)]),   # una sola ora al giorno
                             doc("Verdi", tutte())],
                    timeout=10)
        controlli.controlla(d)
        o = motore.calcola(d)
        s1 = sorted(l.fascia for l in o.lezioni if l.tipo == TIPO_STRUM1)
        self.assertEqual(s1, [fascia(0, 0), fascia(1, 0)])
        controlla_invarianti(self, o)

    def test_laboratorio_lmi_fissato_nel_pomeriggio(self):
        f = fascia(2, 0)  # Mercoledì 13:30
        d = dati_di(studenti=[stud("ROSSI"), stud("NERI", nome="ANNA")],
                    docenti=[doc("Bianchi", tutte(giorni=[0, 1])), doc("Verdi", tutte(giorni=[0, 1])),
                             doc("Galli", [], strumenti="ORCHESTRA")],
                    lmi=[lab("FIATI", "Galli", "Rossi", "Neri")],
                    abbinamenti=[abbinamento("", "FIATI", "Mercoledì", "13:30", "Laboratorio LMI")],
                    timeout=15)
        avvisi = controlli.controlla(d)
        # il docente non era dichiarato disponibile: l'ora viene aperta e segnalata
        self.assertTrue(any("Galli" in a.messaggio and "aperta" in a.messaggio for a in avvisi),
                        [str(a) for a in avvisi])
        o = motore.calcola(d)
        lmi = [l for l in o.lezioni if l.tipo == TIPO_LMI]
        self.assertEqual(len(lmi), 1)
        self.assertEqual((lmi[0].docente, lmi[0].fascia, lmi[0].nome), ("Galli", f, "FIATI"))
        self.assertEqual(set(lmi[0].studenti), {"ROSSI MARIO", "NERI ANNA"})
        # i ragazzi del laboratorio risultano occupati in quell'ora
        for sid in lmi[0].studenti:
            in_quella_fascia = [l for l in o.lezioni_di(sid) if l.fascia == f]
            self.assertEqual(len(in_quella_fascia), 1)
        controlla_invarianti(self, o)

    def test_laboratorio_lmi_inesistente(self):
        with self.assertRaises(ProblemiError) as e:
            controlli.controlla(dati_di(
                studenti=[stud("ROSSI")],
                docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte()), doc("Galli", tutte())],
                lmi=[lab("FIATI", "Galli", "Rossi")],
                abbinamenti=[abbinamento("Galli", "ARCHI", "Mercoledì", "13:30", "Laboratorio LMI")]))
        self.assertIn("ARCHI", str(e.exception))

    def test_laboratorio_lmi_con_un_ragazzo_impegnato(self):
        with self.assertRaises(ProblemiError) as e:
            controlli.controlla(dati_di(
                studenti=[stud("ROSSI"), stud("NERI", nome="ANNA")],
                docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte()), doc("Galli", tutte())],
                lmi=[lab("FIATI", "Galli", "Rossi", "Neri")],
                impegni=[impegno("NERI", nome="ANNA", fasce=[fascia(2, 0)])],
                abbinamenti=[abbinamento("Galli", "FIATI", "Mercoledì", "13:30", "Laboratorio LMI")]))
        self.assertIn("Neri", str(e.exception))


class TestMotoreEsempio(Caso):
    """Test sul dataset reale esempio/orario_2026-27.xlsx (letto, mai scritto)."""

    @classmethod
    def setUpClass(cls):
        if not ESEMPIO.exists():
            raise unittest.SkipTest(f"manca {ESEMPIO}")

    def carica(self, timeout=20) -> DatiInput:
        d = lettura.leggi_dati(ESEMPIO)
        d.parametri.timeout_s = timeout
        controlli.controlla(d)  # risolve i cognomi dei gruppi LMC: il motore ne ha bisogno
        return d

    def test_esempio_completo(self):
        d = lettura.leggi_dati(ESEMPIO)
        d.parametri.timeout_s = 20
        avvisi = controlli.controlla(d)
        self.assertTrue(all(a.livello == "avviso" for a in avvisi))
        self.assertTrue(all(g.studenti for g in d.gruppi))
        o = motore.calcola(d)
        controlla_invarianti(self, o)
        self.assertGreater(len(o.lezioni), 50)
        self.assertIn(o.stato, ("OPTIMAL", "FEASIBLE"))
        TestMotoreEsempio.n_lezioni = len(o.lezioni)

    def test_determinismo(self):
        n = []
        for _ in range(2):
            d = self.carica(timeout=10)
            n.append(len(motore.calcola(d).lezioni))
        self.assertEqual(n[0], n[1])

    def test_ricalcolo_x_rimossa(self):
        d = self.carica()
        o1 = motore.calcola(d)
        # trova una X occupata da una lezione e togliela al docente
        vittima = next(l for l in o1.lezioni if l.tipo != TIPO_ACCOMP)
        d2 = self.carica()
        doc2 = d2.docente(vittima.docente)
        doc2.disponibilita.pop(vittima.fascia, None)
        o2 = motore.calcola(d2, precedente=o1.lezioni)
        controlla_invarianti(self, o2)
        testi = [a.messaggio for a in o2.avvisi]
        riass = [t for t in testi if t.startswith("Rispetto all'orario precedente")]
        self.assertEqual(len(riass), 1, testi[:5])
        # conta gli spostamenti effettivi
        import re
        m = re.search(r"(\d+) lezioni confermate, (\d+) spostate, (-?\d+) nuove", riass[0])
        self.assertIsNotNone(m, riass[0])
        confermate, spostate = int(m.group(1)), int(m.group(2))
        self.assertGreater(confermate, 0)
        self.assertLessEqual(spostate, 5, f"troppe lezioni spostate: {spostate} ({riass[0]})")

    def test_ricalcolo_studente_aggiunto_e_rimosso(self):
        d = self.carica()
        o1 = motore.calcola(d)
        t = lettura.leggi_tabelle(ESEMPIO)
        ts = t[FOGLIO_STUDENTI]
        # rimuovi uno studente di 1ª/2ª (senza LMC) e aggiungine uno nuovo con gli stessi docenti
        i_classe = ts.colonna("Classe")
        vittima = next(r for r in ts.righe if r[i_classe] in ("1", "2"))
        nuovo = list(vittima)
        nuovo[ts.colonna("Cognome")] = "NUOVOARRIVATO"
        nuovo[ts.colonna("Nome")] = "TEST"
        ts.righe = [r for r in ts.righe if r is not vittima] + [nuovo]
        d2 = lettura.costruisci_dati(t, ESEMPIO)
        lettura.applica_trasporti(d2, lettura.leggi_trasporti(ESEMPIO))
        d2.parametri.timeout_s = 20
        controlli.controlla(d2)
        o2 = motore.calcola(d2, precedente=o1.lezioni)
        controlla_invarianti(self, o2)
        riass = next(a.messaggio for a in o2.avvisi if a.messaggio.startswith("Rispetto all'orario precedente"))
        self.assertIn("confermate", riass)
        ids = {s.id for s in d2.studenti}
        self.assertIn("NUOVOARRIVATO TEST", ids)
        self.assertTrue(o2.lezioni_di("NUOVOARRIVATO TEST"))

    def test_salva_e_rilegge_orario_in_copia(self):
        copia = self.dir / "copia.xlsx"
        shutil.copy(ESEMPIO, copia)
        d = lettura.leggi_dati(copia)
        d.parametri.timeout_s = 10
        controlli.controlla(d)
        o = motore.calcola(d)
        lettura.salva_orario_nel_file(copia, o)
        letto = lettura.leggi_orario_precedente(copia)
        self.assertIsNotNone(letto)
        lezioni, quando = letto
        self.assertEqual(len(lezioni), len(o.lezioni))
        self.assertTrue(quando)
        prima = sorted((l.docente, l.fascia, l.tipo, tuple(sorted(l.studenti)), l.due_ore)
                       for l in o.lezioni)
        dopo = sorted((l.docente, l.fascia, l.tipo, tuple(sorted(l.studenti)), l.due_ore)
                      for l in lezioni)
        self.assertEqual(prima, dopo)


# ═════════════════════════════════════════════════════════════════════════════
# 4. Export
# ═════════════════════════════════════════════════════════════════════════════

def _pdftotext(p: Path) -> str:
    r = subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True)
    return r.stdout


def _pagine(p: Path) -> int:
    r = subprocess.run(["pdfinfo", str(p)], capture_output=True, text=True)
    for riga in r.stdout.splitlines():
        if riga.startswith("Pages:"):
            return int(riga.split(":")[1])
    return 0


class TestExport(Caso):

    @classmethod
    def setUpClass(cls):
        if not ESEMPIO.exists():
            raise unittest.SkipTest(f"manca {ESEMPIO}")
        d = lettura.leggi_dati(ESEMPIO)
        d.parametri.timeout_s = 15
        controlli.controlla(d)
        cls.orario = motore.calcola(d)

    def test_file_prodotti_e_fogli(self):
        out = self.dir / "risultati"
        file = esporta_tutto(self.orario, out)
        self.assertEqual([p.name for p in file],
                         ["orario.xlsx", "orario_settimanale.pdf", "orario_docenti.pdf",
                          "orario_studenti.pdf", "orario_settimanale.docx", "orario_docenti.docx",
                          "orario_studenti.docx"])
        for p in file:
            self.assertTrue(p.exists() and p.stat().st_size > 2000, p)
        wb = load_workbook(file[0])
        self.assertEqual(wb.sheetnames, GIORNI + ["Studenti", "Docenti", "Controlli", "LMI"])

    def test_word_leggibile(self):
        """I tre Word si aprono, hanno una pagina per giorno e le celle con lo stile giusto."""
        import docx as _docx

        out = self.dir / "risultati"
        esporta_tutto(self.orario, out)
        doc = _docx.Document(out / "orario_settimanale.docx")
        attese = len(GIORNI) + (1 if self.orario.dati.lmi else 0)
        self.assertEqual(len(doc.sections), attese, "una pagina per giorno, più gli LMI")
        self.assertGreater(doc.sections[0].page_width, doc.sections[0].page_height, "pagina orizzontale")
        t = doc.tables[0]
        self.assertEqual(len(t.rows), 3 + 4)
        self.assertEqual(len(t.columns), 1 + len(self.orario.dati.docenti))
        self.assertEqual([c.text for c in t.rows[0].cells][0], "Docente")
        testo_tabella = "\n".join(c.text for r in t.rows for c in r.cells)
        cognomi = {s.cognome.title() for s in self.orario.dati.studenti}
        self.assertTrue(any(c in testo_tabella for c in cognomi), "nella griglia ci sono i cognomi")
        stili = [(run.bold, run.italic) for r in t.rows[3:] for c in r.cells
                 for p in c.paragraphs for run in p.runs]
        self.assertTrue(any(b for b, _ in stili), "il 1° strumento è in grassetto")
        for nome, verticale in (("orario_docenti.docx", True), ("orario_studenti.docx", True)):
            d2 = _docx.Document(out / nome)
            self.assertTrue(d2.tables, nome)
            self.assertEqual(d2.sections[0].page_width < d2.sections[0].page_height, verticale, nome)

    def test_celle_giorno(self):
        out = self.dir / "risultati"
        esporta_tutto(self.orario, out)
        wb = load_workbook(out / "orario.xlsx")
        dati = self.orario.dati
        colonna = {d.nome: j for j, d in enumerate(dati.docenti, start=2)}
        controllate = 0
        for l in self.orario.lezioni:
            if l.tipo not in (TIPO_STRUM1, TIPO_STRUM2):
                continue
            ws = wb[GIORNI[l.giorno]]
            valore = ws.cell(4 + l.ora, colonna[l.docente]).value or ""
            s = dati.studente(l.studenti[0])
            self.assertIn(s.cognome.title(), valore,
                          f"{s.cognome} atteso in {GIORNI[l.giorno]} {ORE[l.ora]} col. {l.docente}")
            controllate += 1
        self.assertGreater(controllate, 20)
        # celle grigie dove il docente non è disponibile
        grigie = 0
        for g in range(N_GIORNI):
            ws = wb[GIORNI[g]]
            for o in range(N_ORE):
                for d in dati.docenti:
                    c = ws.cell(4 + o, colonna[d.nome])
                    if not d.disponibile(fascia(g, o)):
                        self.assertIsNone(c.value)
                        self.assertEqual(c.fill.fgColor.rgb, "00E7E6E6",
                                         f"{d.nome} {GIORNI[g]} {ORE[o]} non grigia")
                        grigie += 1
        self.assertGreater(grigie, 5)

    def test_foglio_lmi(self):
        out = self.dir / "risultati"
        esporta_tutto(self.orario, out)
        ws = load_workbook(out / "orario.xlsx")["LMI"]
        righe = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
        nomi = {r[0] for r in righe if r[0]}
        attesi = {l.nome for l in self.orario.dati.lmi}
        self.assertTrue(attesi)
        self.assertEqual(nomi, attesi)

    def test_pdf_settimanale(self):
        out = self.dir / "risultati"
        esporta_tutto(self.orario, out)
        p = out / "orario_settimanale.pdf"
        self.assertGreaterEqual(_pagine(p), N_GIORNI)
        testo = _pdftotext(p)
        for g in ("Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"):
            self.assertIn(g, testo)
        self.assertIn("primo strumento", testo)  # legenda
        self.assertIn("secondo strumento", testo)
        cognomi = [s.cognome.title() for s in self.orario.dati.studenti[:8]]
        trovati = [c for c in cognomi if c in testo]
        self.assertGreaterEqual(len(trovati), 6, f"cognomi trovati: {trovati}")

    def test_pdf_docenti_e_studenti(self):
        out = self.dir / "risultati"
        esporta_tutto(self.orario, out)
        dati = self.orario.dati
        pd_ = out / "orario_docenti.pdf"
        self.assertGreaterEqual(_pagine(pd_), 1)
        testo = _pdftotext(pd_)
        for d in dati.docenti[:4]:
            self.assertIn(d.nome, testo)
        ps = out / "orario_studenti.pdf"
        self.assertGreaterEqual(_pagine(ps), 1)
        testo_s = _pdftotext(ps)
        trovati = [s for s in dati.studenti[:8] if s.cognome.title() in testo_s]
        self.assertGreaterEqual(len(trovati), 6)

    def test_cognomi_uguali_con_iniziale(self):
        d = dati_di(studenti=[stud("ROSSI", nome="MARIO"), stud("ROSSI", nome="ANNA")],
                    docenti=[doc("Bianchi", tutte(giorni=[0, 1])), doc("Verdi", tutte(giorni=[0, 1]))],
                    percorso=self.dir / "orario_2026-27.xlsx", timeout=10)
        o = motore.calcola(d)
        out = self.dir / "doppi"
        esporta_tutto(o, out)
        wb = load_workbook(out / "orario.xlsx")
        testi = []
        for g in range(N_GIORNI):
            for r in wb[GIORNI[g]].iter_rows(min_row=4):
                testi += [str(c.value) for c in r if c.value]
        unito = " ".join(testi)
        self.assertIn("Rossi M.", unito)
        self.assertIn("Rossi A.", unito)
        # anno scolastico preso dal nome del file
        self.assertIn("2026-27", _pdftotext(out / "orario_settimanale.pdf"))


# ═════════════════════════════════════════════════════════════════════════════
# 5. Trasporti (senza rete)
# ═════════════════════════════════════════════════════════════════════════════

class TestTrasportiPuro(Caso):

    def test_candidati_geocodifica(self):
        c = trasporti.candidati_geocodifica("Nazario Sauro", "283", "Pistoia")
        self.assertEqual(c[0], "Via Nazario Sauro 283, Pistoia")
        self.assertIn("Via Nazario Sauro, Pistoia", c)
        self.assertEqual(c[-1], "Pistoia")
        c = trasporti.candidati_geocodifica("Via Roma", "", "Montecatini-Terme")
        self.assertEqual(c[0], "Via Roma, Montecatini-Terme")
        self.assertIn("Roma, Montecatini-Terme", c)  # versione "nuda", utile per le frazioni
        c = trasporti.candidati_geocodifica("Località Bellavalle", "68", "Sambuca Pistoiese")
        self.assertEqual(c[0], "Località Bellavalle 68, Sambuca Pistoiese")
        self.assertIn("Bellavalle, Sambuca Pistoiese", c)
        c = trasporti.candidati_geocodifica("Via Cavour", "1", "Sant'Angelo")
        self.assertEqual(c[0], "Via Cavour 1, Sant'Angelo")
        self.assertEqual(trasporti.candidati_geocodifica("", "", "Pistoia"), ["Pistoia"])
        self.assertEqual(trasporti.candidati_geocodifica("", "", ""), [])

    def test_normalizza_comune(self):
        n = trasporti._normalizza_comune
        self.assertEqual(n("Montecatini-Terme"), "montecatini terme")
        self.assertEqual(n("Sant'Angelo"), "sant angelo")
        self.assertEqual(n("Sant’Angelo"), "sant angelo")
        self.assertEqual(n("  SAN   MARCELLO  "), "san marcello")
        self.assertEqual(n("Montecatini-Terme"), n("Montecatini Terme"))

    def test_nel_comune_scarta_provincia(self):
        f = trasporti._nel_comune
        self.assertTrue(f({"address": {"town": "Montale"}}, "Montale"))
        self.assertTrue(f({"address": {"village": "Montecatini Terme"}}, "Montecatini-Terme"))
        # la provincia si chiama come il comune: non deve bastare
        self.assertFalse(f({"address": {"county": "Pistoia", "town": "Montale"}}, "Pistoia"))
        self.assertTrue(f({"address": {}}, ""))  # comune non indicato → tutto passa

    def test_ora_italiana_riserva(self):
        tz = trasporti._OraItaliana()
        self.assertEqual(tz._ultima_domenica(2026, 3).date(), date(2026, 3, 29))
        self.assertEqual(tz._ultima_domenica(2026, 10).date(), date(2026, 10, 25))
        self.assertEqual(tz._ultima_domenica(2025, 3).date(), date(2025, 3, 30))
        self.assertEqual(tz._ultima_domenica(2025, 12).date(), date(2025, 12, 28))
        ore = lambda dt: tz.utcoffset(dt).total_seconds() / 3600
        self.assertEqual(ore(datetime(2026, 1, 15, 12)), 1)
        self.assertEqual(ore(datetime(2026, 7, 15, 12)), 2)
        # confini
        self.assertEqual(ore(datetime(2026, 3, 29, 1, 59)), 1)
        self.assertEqual(ore(datetime(2026, 3, 29, 2, 0)), 2)
        self.assertEqual(ore(datetime(2026, 10, 25, 2, 59)), 2)
        self.assertEqual(ore(datetime(2026, 10, 25, 3, 0)), 1)
        self.assertEqual(tz.tzname(datetime(2026, 7, 1)), "CEST")
        self.assertEqual(tz.tzname(datetime(2026, 1, 1)), "CET")

    def test_prossimo_martedi(self):
        d = trasporti.prossimo_martedi(date(2026, 9, 9))  # mercoledì
        self.assertEqual(d.weekday(), 1)
        self.assertGreaterEqual((d - date(2026, 9, 9)).days, 3)


SCUOLA = Coordinate(43.93, 10.91, "scuola")
CASA = Coordinate(44.0, 10.8, "casa")


def _iso(dt: datetime) -> str:
    return dt.astimezone(trasporti.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _itin(giorno: date, ora_p: int, min_p: int, durata: int, giorni_dopo: int = 0,
          modo="BUS", nome="56") -> dict:
    inizio = datetime(giorno.year, giorno.month, giorno.day, ora_p, min_p, tzinfo=trasporti.FUSO)
    fine = inizio + timedelta(minutes=durata, days=giorni_dopo)
    return {"startTime": _iso(inizio), "endTime": _iso(fine),
            "legs": [{"mode": "WALK"}, {"mode": modo, "routeShortName": nome}]}


class TestRitorni(Caso):

    def patch_get(self, risposta):
        vero = trasporti._get
        self.chiamate = []

        def finto(url, timeout=60):
            self.chiamate.append(url)
            return risposta(url) if callable(risposta) else risposta

        trasporti._get = finto
        self.addCleanup(lambda: setattr(trasporti, "_get", vero))

    def test_itinerario_normale(self):
        g = date(2026, 9, 15)
        itin = [_itin(g, 14, 40, 30), _itin(g, 15, 40, 30), _itin(g, 16, 40, 30), _itin(g, 17, 40, 30)]
        self.patch_get({"itineraries": itin, "direct": []})
        minuti, arrivi, mezzi = trasporti.ritorni(SCUOLA, CASA, g)
        self.assertEqual(minuti, [40, 40, 40, 40])  # 5 (attesa fine lezione) + 5 + 30
        self.assertEqual(arrivi, ["15:10", "16:10", "17:10", "18:10"])
        self.assertTrue(all(m == "bus 56" for m in mezzi), mezzi)

    def test_arrivo_giorno_dopo_scartato(self):
        g = date(2026, 9, 15)
        self.patch_get({"itineraries": [_itin(g, 14, 40, 30, giorni_dopo=1)]})
        minuti, arrivi, mezzi = trasporti.ritorni(SCUOLA, CASA, g)
        self.assertEqual(minuti, [None] * 4)
        self.assertEqual(arrivi, [""] * 4)

    def test_percorso_solo_a_piedi_in_direct(self):
        g = date(2026, 9, 15)
        piedi = {"startTime": _iso(datetime(2026, 9, 15, 14, 40, tzinfo=trasporti.FUSO)),
                 "endTime": _iso(datetime(2026, 9, 15, 15, 5, tzinfo=trasporti.FUSO)),
                 "legs": [{"mode": "WALK"}]}
        self.patch_get({"itineraries": [], "direct": [piedi]})
        minuti, arrivi, mezzi = trasporti.ritorni(SCUOLA, CASA, g)
        self.assertEqual(minuti[0], 35)
        self.assertEqual(mezzi[0], "a piedi")

    def test_nessun_itinerario(self):
        self.patch_get({"itineraries": [], "direct": []})
        minuti, _, _ = trasporti.ritorni(SCUOLA, CASA, date(2026, 9, 15))
        self.assertEqual(minuti, [None] * 4)

    def test_troppo_tardi(self):
        g = date(2026, 9, 15)
        # parte alle 14:40, arriva alle 20:10 → 335 minuti > MINUTI_SENZA_MEZZO
        self.patch_get({"itineraries": [_itin(g, 14, 40, 330)]})
        minuti, arrivi, mezzi = trasporti.ritorni(SCUOLA, CASA, g)
        self.assertIsNone(minuti[0])
        self.assertIn("troppo tardi", mezzi[0])
        self.assertEqual(arrivi[0], "20:10")
        self.assertGreater(MINUTI_SENZA_MEZZO, 0)
        # la stessa corsa vista dalla 10ª fascia (17:35) rientra nei limiti
        self.assertIsNone(minuti[3])  # nessun itinerario dopo le 17:35

    def test_errore_http(self):
        def alza(url, timeout=60):
            raise urllib.error.HTTPError(url, 500, "boom", {}, None)

        self.patch_get(alza)
        minuti, _, _ = trasporti.ritorni(SCUOLA, CASA, date(2026, 9, 15))
        self.assertEqual(minuti, [None] * 4)

    def test_itinerari_url(self):
        self.patch_get({"itineraries": []})
        trasporti._itinerari(SCUOLA, CASA, datetime(2026, 9, 15, 14, 35, tzinfo=trasporti.FUSO))
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.chiamate[0]).query)
        self.assertEqual(q["fromPlace"][0], "43.93,10.91")
        self.assertEqual(q["toPlace"][0], "44.0,10.8")
        self.assertEqual(q["arriveBy"][0], "false")
        self.assertEqual(q["maxPostTransitTime"][0], str(trasporti.CAMMINO_MAX_DALLA_FERMATA))
        self.assertTrue(q["time"][0].endswith("Z"))


class TestTrasportiFile(Caso):

    def test_salva_leggi_applica(self):
        p = self.nuovo_file()
        t = lettura.leggi_tabelle(p)
        t[FOGLIO_STUDENTI].righe = [
            _riga(COLONNE_STUDENTI, stud("ROSSI", comune="Pistoia", indirizzo="Via Roma", civico="1")),
            _riga(COLONNE_STUDENTI, stud("NERI", nome="ANNA", comune="Pistoia",
                                         indirizzo="Via Verdi", civico="2")),
        ]
        t[FOGLIO_DOCENTI].righe = [_riga(COLONNE_DOCENTI, doc("Bianchi", tutte())),
                                   _riga(COLONNE_DOCENTI, doc("Verdi", tutte()))]
        t[FOGLIO_GRUPPI].righe = []
        t[FOGLIO_LMI].righe = []
        lettura.salva_tabelle(p, t)
        dati = lettura.leggi_dati(p)
        rossi, neri = dati.studenti[0], dati.studenti[1]
        tr = {
            rossi.id: Trasporto([12, 20, None, 44], ["14:47", "15:55", "", "18:19"],
                                ["bus 10", "bus 10", "", "treno + bus"],
                                43.9, 10.9, rossi.indirizzo_completo, "OK", "09/09/2026 10:00"),
            # indirizzo diverso da quello attuale: non deve essere applicato
            neri.id: Trasporto([30, 30, 30, 30], ["", "", "", ""], ["bus", "bus", "bus", "bus"],
                               43.8, 10.8, "Via Vecchia 9, Pistoia", "OK", "09/09/2026 10:00"),
        }
        lettura.salva_trasporti(p, tr)
        self.assertIn(FOGLIO_TRASPORTI, load_workbook(p).sheetnames)
        riletti = lettura.leggi_trasporti(p)
        self.assertEqual(set(riletti), set(tr))
        a = riletti[rossi.id]
        self.assertEqual(a.minuti, [12, 20, None, 44])
        self.assertEqual(a.arrivi, ["14:47", "15:55", "", "18:19"])
        self.assertEqual(a.mezzi, ["bus 10", "bus 10", "", "treno + bus"])
        self.assertEqual((a.lat, a.lon), (43.9, 10.9))
        self.assertEqual(a.indirizzo_usato, rossi.indirizzo_completo)
        self.assertEqual(a.esito, "OK")
        self.assertEqual(a.aggiornato, "09/09/2026 10:00")
        self.assertEqual(a.minuti_min, 12)
        self.assertTrue(a.valido)
        n = lettura.applica_trasporti(dati, riletti)
        self.assertEqual(n, 1)
        self.assertIsNotNone(rossi.trasporto)
        self.assertIsNone(neri.trasporto, "trasporto con indirizzo diverso applicato per errore")
        # leggi_dati aggancia i trasporti da solo
        dati2 = lettura.leggi_dati(p)
        self.assertIsNotNone(dati2.studenti[0].trasporto)
        self.assertIsNone(dati2.studenti[1].trasporto)

    def test_leggi_trasporti_senza_foglio(self):
        p = self.nuovo_file()
        self.assertEqual(lettura.leggi_trasporti(p), {})


class TestAggiornaTrasporti(Caso):

    def setUp(self):
        super().setUp()
        self.chiamate = []
        vero = trasporti._get
        self.addCleanup(lambda: setattr(trasporti, "_get", vero))
        trasporti._get = self._finto
        trasporti.PAUSA_NOMINATIM = 0.0

    def _finto(self, url, timeout=60):
        self.chiamate.append(url)
        if "photon" in url:
            # il servizio vero restituisce la via chiesta: il programma scarta le vie diverse
            chiesto = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)["q"][0]
            via = re.sub(r"\s*\d+\S*\s*$", "", chiesto.split(",")[0]).strip() or "Via Roma"
            return {"features": [{"geometry": {"coordinates": [10.9, 43.9]},
                                  "properties": {"name": via, "city": "Pistoia",
                                                 "osm_value": "residential"}}]}
        if "nominatim" in url:
            return [{"lat": "43.93", "lon": "10.91", "display_name": "Pistoia",
                     "addresstype": "town", "address": {"town": "Pistoia"}}]
        if "transitous" in url:
            q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
            quando = datetime.fromisoformat(q["time"][0].replace("Z", "+00:00")).astimezone(trasporti.FUSO)
            g = quando.date()
            return {"itineraries": [_itin(g, h, 40, 30) for h in (14, 15, 16, 17, 18)], "direct": []}
        raise AssertionError(f"URL inatteso: {url}")

    def _dati(self, n=3):
        return dati_di(
            studenti=[stud(f"STUD{i}", nome=f"N{i}", comune="Pistoia",
                           indirizzo="Via Roma", civico=str(i)) for i in range(n)],
            docenti=[doc("Bianchi", tutte()), doc("Verdi", tutte())],
            parametri={"Indirizzo della scuola": "Via Cavour 1, Pistoia"})

    def test_aggiorna_e_riepilogo(self):
        d = self._dati()
        r = trasporti.aggiorna_trasporti(d)
        self.assertEqual(len(r), 3)
        self.assertTrue(all(t.esito == "OK" for t in r.values()), {k: v.esito for k, v in r.items()})
        for s in d.studenti:
            t = r[s.id]
            self.assertEqual(t.minuti, [40, 40, 40, 40])
            self.assertEqual(t.indirizzo_usato, s.indirizzo_completo)
            self.assertEqual((t.lat, t.lon), (43.9, 10.9))
        testo = trasporti.riepilogo(r, d)
        self.assertIn("Trasporti calcolati per 3 studenti", testo)

    def test_senza_indirizzo_scuola(self):
        d = self._dati()
        d.parametri.indirizzo_scuola = "  "
        with self.assertRaises(TrasportiError) as cm:
            trasporti.aggiorna_trasporti(d)
        self.assertIn("Manca l'indirizzo della scuola", str(cm.exception))

    def test_solo_nuovi_non_richiede(self):
        d = self._dati()
        precedenti = {s.id: Trasporto([40] * 4, ["15:10"] * 4, ["bus 56"] * 4, 43.9, 10.9,
                                      s.indirizzo_completo, "OK", "ieri") for s in d.studenti}
        self.chiamate.clear()
        r = trasporti.aggiorna_trasporti(d, precedenti=precedenti, solo_nuovi=True)
        self.assertEqual(len(r), 3)
        self.assertTrue(all("transitous" not in u for u in self.chiamate),
                        [u[:40] for u in self.chiamate])
        self.assertEqual(r[d.studenti[0].id].aggiornato, "ieri")
        # con uno studente nuovo si fa una sola richiesta di percorsi
        d2 = self._dati(4)
        self.chiamate.clear()
        r2 = trasporti.aggiorna_trasporti(d2, precedenti=precedenti, solo_nuovi=True)
        self.assertEqual(len(r2), 4)
        self.assertEqual(sum(1 for u in self.chiamate if "transitous" in u), 1,
                         [u[:40] for u in self.chiamate])

    def test_annulla(self):
        d = self._dati()
        stop = threading.Event()
        stop.set()
        r = trasporti.aggiorna_trasporti(d, annulla=stop)
        self.assertEqual(r, {})

    def test_riuso_coordinate(self):
        d = self._dati(1)
        s = d.studenti[0]
        precedenti = {s.id: Trasporto([40] * 4, [""] * 4, [""] * 4, 43.5, 10.5,
                                      s.indirizzo_completo, "OK", "ieri")}
        self.chiamate.clear()
        trasporti.aggiorna_trasporti(d, precedenti=precedenti)
        # nessuna geocodifica per lo studente: solo quella della scuola
        self.assertEqual(sum(1 for u in self.chiamate if "photon" in u), 1,
                         [u[:40] for u in self.chiamate])
        q = urllib.parse.parse_qs(urllib.parse.urlparse(
            next(u for u in self.chiamate if "transitous" in u)).query)
        self.assertEqual(q["toPlace"][0], "43.5,10.5")


# ═════════════════════════════════════════════════════════════════════════════

class Riepilogo(unittest.TextTestResult):
    pass


def main() -> int:
    caricatore = unittest.TestLoader()
    caricatore.sortTestMethodsUsing = None
    suite = caricatore.loadTestsFromModule(sys.modules[__name__])
    risultato = unittest.TextTestRunner(verbosity=2, resultclass=Riepilogo).run(suite)
    tot = risultato.testsRun
    ko = len(risultato.failures) + len(risultato.errors)
    print("\n" + "=" * 74)
    print(f"RIEPILOGO: {tot} test eseguiti — {tot - ko - len(risultato.skipped)} passati, "
          f"{len(risultato.failures)} falliti, {len(risultato.errors)} errori, "
          f"{len(risultato.skipped)} saltati")
    if ko:
        print("\nDa correggere:")
        for caso, _ in risultato.failures + risultato.errors:
            print(f"  ✗ {caso.id().split('.', 1)[-1]}")
    else:
        print("Tutto a posto.")
    print("=" * 74)
    return 0 if ko == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
