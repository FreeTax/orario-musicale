"""Tempi di ritorno a casa con i mezzi pubblici, per ogni studente e per ogni fascia.

Servizi usati (gratuiti, open source, senza chiave):
  * Nominatim (OpenStreetMap) per trasformare un indirizzo in coordinate  — max 1 richiesta al secondo
  * Transitous (api.transitous.org, motore MOTIS) per i percorsi con bus e treni

Il risultato viene salvato nel foglio "Trasporti" del file di input (vedi lettura.salva_trasporti),
così il calcolo dell'orario resta offline.
"""

from __future__ import annotations

import json
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Callable
from datetime import timezone, tzinfo
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .costanti import MINUTI_SENZA_MEZZO, ORE_FINE, ORE_PARTENZA
from .modello import DatiInput, Studente, Trasporto

USER_AGENT = "OrarioLiceoMusicale/1.0 (programma scolastico; contatto: francesco.mazzola@kinoa.studio)"
NOMINATIM = "https://nominatim.openstreetmap.org/search"
PHOTON = "https://photon.komoot.io/api/"
CENTRO_RICERCA = (43.93, 10.91)  # Pistoia: aiuta Photon a preferire risultati vicini
STUDENTI_IN_PARALLELO = 2
CAMMINO_MAX_DALLA_FERMATA = 1800  # secondi a piedi dall'ultima fermata a casa (30 min): case di campagna  # ogni studente fa 4 richieste a Transitous insieme → al massimo 8 contemporanee
TRANSITOUS = "https://api.transitous.org/api/v1/plan"
UTC = timezone.utc


class _OraItaliana(tzinfo):
    """Riserva se manca il database dei fusi (Windows senza il pacchetto tzdata).

    Regola europea: ora legale (+2) dall'ultima domenica di marzo all'ultima di ottobre, altrimenti +1.
    """

    @staticmethod
    def _ultima_domenica(anno: int, mese: int) -> datetime:
        d = datetime(anno + (mese == 12), mese % 12 + 1, 1) - timedelta(days=1)
        return d - timedelta(days=(d.weekday() + 1) % 7)

    def _legale(self, dt: datetime) -> bool:
        naive = dt.replace(tzinfo=None)
        inizio = self._ultima_domenica(naive.year, 3).replace(hour=2)
        fine = self._ultima_domenica(naive.year, 10).replace(hour=3)
        return inizio <= naive < fine

    def utcoffset(self, dt):
        return timedelta(hours=2 if dt and self._legale(dt) else 1)

    def dst(self, dt):
        return timedelta(hours=1 if dt and self._legale(dt) else 0)

    def tzname(self, dt):
        return "CEST" if dt and self._legale(dt) else "CET"


try:
    FUSO: tzinfo = ZoneInfo("Europe/Rome")
except (ZoneInfoNotFoundError, KeyError):  # Windows senza tzdata
    FUSO = _OraItaliana()
PAUSA_NOMINATIM = 1.1  # secondi tra due richieste (regola d'uso di Nominatim)
PAUSA_TRANSITOUS = 0.4
TENTATIVI = 3


class TrasportiError(Exception):
    pass


@dataclass
class Coordinate:
    lat: float
    lon: float
    descrizione: str


def _get(url: str, timeout: int = 60) -> dict | list:
    ultimo: Exception | None = None
    for tentativo in range(TENTATIVI):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            ultimo = e
            if e.code in (400, 404):
                raise
            time.sleep(2 * (tentativo + 1))
        except Exception as e:  # rete assente, timeout…
            ultimo = e
            time.sleep(2 * (tentativo + 1))
    raise TrasportiError(f"Servizio non raggiungibile: {ultimo}")


# ── Geocodifica ──────────────────────────────────────────────────────────────

TIPI_VIA = ("via", "viale", "v.le", "piazza", "p.za", "p.zza", "corso", "largo", "vicolo", "borgo", "strada",
            "località", "localita", "loc.", "loc", "frazione", "fraz.", "piazzale", "lungarno", "traversa", "podere")


def _normalizza_comune(c: str) -> str:
    # 'Montecatini-Terme' nei recapiti, 'Montecatini Terme' nelle mappe: trattini e apostrofi diventano spazi
    return " ".join(c.lower().replace("-", " ").replace("'", " ").replace("’", " ").split())


def candidati_geocodifica(indirizzo: str, civico: str, comune: str) -> list[str]:
    """Stringhe da cercare, dalla più precisa alla più generica."""
    via = " ".join(indirizzo.split())
    out: list[str] = []
    if via:
        primo = via.split()[0].lower().rstrip(".'’")
        ha_tipo = primo in [t.rstrip(".") for t in TIPI_VIA]
        via_ok = via if ha_tipo else f"Via {via}"
        if civico:
            out.append(f"{via_ok} {civico}, {comune}" if comune else f"{via_ok} {civico}")
        out.append(f"{via_ok}, {comune}" if comune else via_ok)
        # solo il nome, senza 'Via'/'Località' (utile per frazioni: 'Bellavalle, Sambuca Pistoiese')
        nudo = " ".join(via.split()[1:]) if ha_tipo else via
        if nudo and nudo.lower() != via_ok.lower():
            out.append(f"{nudo}, {comune}" if comune else nudo)
    if comune:
        out.append(comune)
    return out


_CAMPI_COMUNE = ("city", "town", "village", "hamlet", "municipality", "suburb", "isolated_dwelling", "locality")
_TIPI_AMMINISTRATIVI = {"administrative", "municipality", "city", "town", "county", "state", "province"}


def _nel_comune(r: dict, comune: str) -> bool:
    if not comune:
        return True
    c = _normalizza_comune(comune)
    addr = r.get("address") or {}
    return any(_normalizza_comune(str(addr.get(k, ""))) == c for k in _CAMPI_COMUNE)


_lucchetto_nominatim = threading.Lock()
_ultima_nominatim = [0.0]


def _nominatim(q: str) -> list:
    """Una richiesta a Nominatim rispettando la regola di 1 al secondo, anche da più thread."""
    with _lucchetto_nominatim:
        attesa = PAUSA_NOMINATIM - (time.time() - _ultima_nominatim[0])
        if attesa > 0:
            time.sleep(attesa)
        url = NOMINATIM + "?" + urllib.parse.urlencode({"q": q, "format": "json", "limit": 3, "countrycodes": "it",
                                                          "addressdetails": 1})
        try:
            return _get(url, timeout=30) or []
        finally:
            _ultima_nominatim[0] = time.time()


def _nel_comune_photon(p: dict, comune: str) -> bool:
    if not comune:
        return True
    c = _normalizza_comune(comune)
    # niente "county": la provincia di Pistoia si chiama come il comune e farebbe passare Montale o Quarrata
    return any(_normalizza_comune(str(p.get(k, ""))) == c for k in ("city", "town", "village", "municipality",
                                                                    "district", "locality"))


def geocodifica_photon(indirizzo: str, comune: str = "", civico: str = "") -> list[Coordinate]:
    """Photon (OpenStreetMap): veloce e tollerante con le abbreviazioni. Accetta solo risultati nel comune."""
    risultati: list[Coordinate] = []
    for q in candidati_geocodifica(indirizzo, civico, comune)[:3]:  # senza il solo comune: quello lo fa Nominatim
        if comune and q.lower() == comune.lower():
            continue
        url = PHOTON + "?" + urllib.parse.urlencode({"q": q, "limit": 5, "lat": CENTRO_RICERCA[0], "lon": CENTRO_RICERCA[1]})
        try:
            d = _get(url, timeout=30)
        except (TrasportiError, urllib.error.HTTPError):
            return risultati
        for f in (d.get("features", []) if isinstance(d, dict) else []):
            pr = f.get("properties", {})
            if not _nel_comune_photon(pr, comune):
                continue
            if pr.get("osm_value") in ("administrative", "city", "town", "municipality"):
                continue
            lon, lat = f["geometry"]["coordinates"]
            c = Coordinate(float(lat), float(lon), f"{pr.get('name') or pr.get('street', '')} {pr.get('housenumber', '')}, "
                                                    f"{pr.get('city') or pr.get('town') or pr.get('village') or comune}".strip())
            if all(abs(c.lat - x.lat) > 1e-4 or abs(c.lon - x.lon) > 1e-4 for x in risultati):
                risultati.append(c)
            break
        if risultati:
            break  # il primo candidato che dà un risultato nel comune basta
    return risultati


def geocodifica(indirizzo: str, comune: str = "", civico: str = "") -> list[Coordinate]:
    """Coordinate candidate, dalla più precisa alla più generica: prima Photon, poi Nominatim come riserva."""
    risultati: list[Coordinate] = geocodifica_photon(indirizzo, comune, civico)
    if risultati:
        return risultati  # Photon basta: Nominatim (1 richiesta/secondo) solo se non ha trovato nulla
    visti = set()
    for q in candidati_geocodifica(indirizzo, civico, comune):
        if q.lower() in visti:
            continue
        visti.add(q.lower())
        risposta = _nominatim(q)
        solo_comune = comune and q.lower() == comune.lower()
        for r in risposta or []:
            nome = r.get("display_name", "")
            if not _nel_comune(r, comune):
                continue  # trovato in un altro comune (o solo nella stessa provincia): non è lui
            if not solo_comune and r.get("addresstype") in _TIPI_AMMINISTRATIVI:
                continue  # cercavo una via e mi ha dato il comune intero: lo tengo solo come ultimo ripiego
            c = Coordinate(float(r["lat"]), float(r["lon"]), nome[:80])
            if all(abs(c.lat - x.lat) > 1e-4 or abs(c.lon - x.lon) > 1e-4 for x in risultati):
                risultati.append(c)
            break
    return risultati


# ── Percorsi ─────────────────────────────────────────────────────────────────

def prossimo_martedi(oggi: date | None = None) -> date:
    """Il prossimo martedì ad almeno 3 giorni da oggi (gli orari dei mezzi sono pubblicati con anticipo)."""
    d = (oggi or date.today()) + timedelta(days=3)
    while d.weekday() != 1:
        d += timedelta(days=1)
    return d


def _descrivi(itinerario: dict) -> str:
    pezzi = []
    for l in itinerario.get("legs", []):
        if l.get("mode") == "WALK":
            continue
        nome = l.get("routeShortName") or l.get("mode", "")
        pezzi.append(f"{'treno' if 'RAIL' in str(l.get('mode')) else 'bus'} {nome}".strip())
    return " + ".join(pezzi) if pezzi else "a piedi"


def _itinerari(da: Coordinate, a: Coordinate, quando: datetime, quanti: int = 40) -> list[dict]:
    url = TRANSITOUS + "?" + urllib.parse.urlencode({
        "fromPlace": f"{da.lat},{da.lon}", "toPlace": f"{a.lat},{a.lon}",
        "time": quando.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "arriveBy": "false", "numItineraries": quanti,
        "maxPostTransitTime": CAMMINO_MAX_DALLA_FERMATA,
    })
    try:
        d = _get(url)
    except urllib.error.HTTPError:
        return []
    if not isinstance(d, dict):
        return []
    return list(d.get("itineraries", [])) + list(d.get("direct", []))


def _ora_locale(ts: str) -> datetime:
    return datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(FUSO)


def ritorni(da: Coordinate, a: Coordinate, giorno: date) -> tuple[list[int | None], list[str], list[str]]:
    """Per ciascuna delle 4 fasce: (minuti da fine lezione ad arrivo a casa, ora di arrivo, mezzi).

    Una sola richiesta a Transitous (tutte le partenze del pomeriggio), una seconda solo se la prima
    non copre l'ultima fascia. Transitous limita il ritmo a circa una richiesta ogni 3 secondi per
    utente: meno richieste = aggiornamento più veloce.
    """
    partenze = []
    for ora in ORE_PARTENZA:
        hh, mm = map(int, ora.split(":"))
        partenze.append(datetime(giorno.year, giorno.month, giorno.day, hh, mm, tzinfo=FUSO))
    itin = _itinerari(da, a, partenze[0])
    inizi = [_ora_locale(i["startTime"]) for i in itin if "startTime" in i]
    if itin and inizi and max(inizi) < partenze[-1] + timedelta(minutes=20):
        itin += _itinerari(da, a, partenze[-1], quanti=10)
    minuti: list[int | None] = []
    arrivi: list[str] = []
    mezzi: list[str] = []
    for partenza in partenze:
        migliore = None
        for it in itin:
            try:
                inizio, fine = _ora_locale(it["startTime"]), _ora_locale(it["endTime"])
            except Exception:
                continue
            if inizio < partenza or fine.date() != giorno:
                continue
            if migliore is None or fine < migliore[0]:
                migliore = (fine, it)
        if migliore is None:
            minuti.append(None)
            arrivi.append("")
            mezzi.append("")
            continue
        fine, it = migliore
        m = int((fine - partenza).total_seconds() // 60) + 5  # +5: da fine lezione alla partenza
        if m > MINUTI_SENZA_MEZZO:
            minuti.append(None)
            arrivi.append(fine.strftime("%H:%M"))
            mezzi.append(_descrivi(it) + " (troppo tardi)")
        else:
            minuti.append(m)
            arrivi.append(fine.strftime("%H:%M"))
            mezzi.append(_descrivi(it))
    return minuti, arrivi, mezzi


# ── Aggiornamento completo ───────────────────────────────────────────────────

def aggiorna_trasporti(dati: DatiInput, progresso: Callable[[str], None] | None = None,
                       annulla: threading.Event | None = None,
                       precedenti: dict[str, Trasporto] | None = None,
                       solo_nuovi: bool = False) -> dict[str, Trasporto]:
    """Calcola i tempi di ritorno per tutti gli studenti con un indirizzo. Ritorna id → Trasporto.

    `precedenti`: risultati già presenti nel file; le coordinate già note si riusano (meno richieste).
    `solo_nuovi`: chi ha già un risultato OK con lo stesso indirizzo viene tenuto com'è, senza richieste.
    Studenti senza indirizzo né comune vengono saltati (restano ai km).
    """
    def dire(m: str) -> None:
        if progresso:
            progresso(m)

    if not dati.parametri.indirizzo_scuola.strip():
        raise TrasportiError("Manca l'indirizzo della scuola nel foglio Parametri (riga 'Indirizzo della scuola').")
    dire("Cerco le coordinate della scuola…")
    parti = [x.strip() for x in dati.parametri.indirizzo_scuola.split(",")]
    trovati = geocodifica(", ".join(parti[:-1]) if len(parti) > 1 else parti[0], parti[-1] if len(parti) > 1 else "")
    if not trovati:
        raise TrasportiError(f"Indirizzo della scuola non trovato: '{dati.parametri.indirizzo_scuola}'. "
                             "Scriverlo come 'Via e numero, Comune'.")
    scuola = trovati[0]
    giorno = prossimo_martedi()
    adesso = datetime.now().strftime("%d/%m/%Y %H:%M")
    risultati: dict[str, Trasporto] = {}
    da_fare = [s for s in dati.studenti if s.indirizzo_completo]
    if solo_nuovi and precedenti:
        tenuti = [s for s in da_fare if (p := precedenti.get(s.id)) and p.esito == "OK"
                  and p.indirizzo_usato == s.indirizzo_completo]
        for s in tenuti:
            risultati[s.id] = precedenti[s.id]
        da_fare = [s for s in da_fare if s.id not in risultati]
        dire(f"{len(tenuti)} studenti già calcolati con lo stesso indirizzo: tenuti. Da fare: {len(da_fare)}.")
    lucchetto = threading.Lock()
    fatti = [0]

    def uno(s: Studente) -> None:
        if annulla is not None and annulla.is_set():
            return
        prec = (precedenti or {}).get(s.id)
        candidati: list[Coordinate] = []
        if prec and prec.lat is not None and prec.indirizzo_usato == s.indirizzo_completo and prec.esito == "OK":
            candidati = [Coordinate(prec.lat, prec.lon, prec.indirizzo_usato)]
        else:
            try:
                candidati = geocodifica(s.indirizzo, s.comune, s.civico)
            except TrasportiError as e:
                with lucchetto:
                    risultati[s.id] = Trasporto([None] * 4, esito=f"errore: {e}", aggiornato=adesso,
                                                indirizzo_usato=s.indirizzo_completo)
                return
        if not candidati:
            with lucchetto:
                risultati[s.id] = Trasporto([None] * 4, esito="indirizzo non trovato", aggiornato=adesso,
                                            indirizzo_usato=s.indirizzo_completo)
            return
        esito, scelto = "nessun percorso trovato", candidati[0]
        minuti: list[int | None] = [None] * 4
        arrivi, mezzi = [""] * 4, [""] * 4
        if s.comune:  # ultimo ripiego: il centro del comune (Nominatim, 1 richiesta/secondo)
            candidati = candidati + [None]  # type: ignore[list-item]
        for coord in candidati:  # dal punto più preciso al più generico, finché si trova un mezzo
            if annulla is not None and annulla.is_set():
                return
            if coord is None:
                try:
                    r = next((x for x in _nominatim(s.comune) if _nel_comune(x, s.comune)), None)
                except TrasportiError:
                    r = None
                if r is None:
                    continue
                coord = Coordinate(float(r["lat"]), float(r["lon"]), f"centro di {s.comune}")
            try:
                minuti, arrivi, mezzi = ritorni(scuola, coord, giorno)
            except TrasportiError as e:
                esito = "errore: " + str(e)
                continue
            if any(m is not None for m in minuti):
                esito, scelto = "OK", coord
                break
        with lucchetto:
            risultati[s.id] = Trasporto(minuti, arrivi, mezzi, scelto.lat, scelto.lon, s.indirizzo_completo, esito, adesso)
            fatti[0] += 1
            dire(f"{fatti[0]}/{len(da_fare)}  {s.cognome.title()} {s.nome.title()} ({s.comune or s.indirizzo}): "
                 f"{'ok' if esito == 'OK' else esito}")

    from concurrent.futures import ThreadPoolExecutor
    dire(f"Cerco indirizzi e mezzi per {len(da_fare)} studenti…")
    with ThreadPoolExecutor(max_workers=STUDENTI_IN_PARALLELO) as ex:
        list(ex.map(uno, da_fare))
    if annulla is not None and annulla.is_set():
        dire("Interrotto dall'utente.")
    return risultati


def riepilogo(risultati: dict[str, Trasporto], dati: DatiInput) -> str:
    ok = sum(1 for t in risultati.values() if t.esito == "OK")
    non_trovati = [i for i, t in risultati.items() if t.esito == "indirizzo non trovato"]
    senza = [i for i, t in risultati.items() if t.esito.startswith("nessun")]
    errori = [i for i, t in risultati.items() if t.esito.startswith("errore")]
    senza_indirizzo = [s for s in dati.studenti if not s.indirizzo_completo]
    righe = [f"Trasporti calcolati per {ok} studenti."]
    if non_trovati:
        righe.append(f"Indirizzo non trovato per {len(non_trovati)}: " + ", ".join(i.title() for i in non_trovati[:8])
                     + (" …" if len(non_trovati) > 8 else "") + ". Controllare via e comune nel foglio Studenti.")
    if senza:
        righe.append(f"Nessun mezzo trovato per {len(senza)}: " + ", ".join(i.title() for i in senza[:8])
                     + (" …" if len(senza) > 8 else "") + ". Restano ai km.")
    if errori:
        righe.append(f"Errori di rete per {len(errori)} studenti: riprovare più tardi.")
    if senza_indirizzo:
        righe.append(f"{len(senza_indirizzo)} studenti senza indirizzo nel foglio: restano ai km.")
    return "\n".join(righe)


__all__ = ["aggiorna_trasporti", "geocodifica", "ritorni", "riepilogo", "TrasportiError", "ORE_FINE"]
