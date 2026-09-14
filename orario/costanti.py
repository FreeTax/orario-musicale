"""Costanti condivise: giorni, fasce orarie, ore per classe."""

GIORNI = ["Lun", "Mar", "Mer", "Gio", "Ven"]
GIORNI_LUNGHI = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]
ORE = ["13:30", "14:30", "15:30", "16:30"]
ORE_FINE = ["14:30", "15:30", "16:30", "17:30"]
ORE_LABEL = ["7ª", "8ª", "9ª", "10ª"]

N_GIORNI = len(GIORNI)
N_ORE = len(ORE)
N_FASCE = N_GIORNI * N_ORE  # 20

# Nomi delle 20 colonne di disponibilità nel foglio Docenti: "Lun 13:30" ... "Ven 16:30"
FASCE = [f"{g} {o}" for g in GIORNI for o in ORE]

# Una aula per giorno: il docente può cambiare stanza da un giorno all'altro
COLONNE_AULE = [f"Aula {g}" for g in GIORNI]


def fascia(giorno: int, ora: int) -> int:
    return giorno * N_ORE + ora


def giorno_ora(f: int) -> tuple[int, int]:
    return divmod(f, N_ORE)


def nome_fascia(f: int) -> str:
    g, o = giorno_ora(f)
    return f"{GIORNI_LUNGHI[g]} {ORE[o]}-{ORE_FINE[o]}"


# Ore settimanali per classe: (1° strumento, 2° strumento, musica da camera LMC)
ORE_PER_CLASSE: dict[int, tuple[int, int, int]] = {
    1: (2, 1, 0),
    2: (2, 1, 0),
    3: (1, 1, 1),
    4: (1, 1, 1),
    5: (2, 0, 1),
}

CLASSI_LMC = (3, 4, 5)
MIN_GRUPPO_LMC = 2
MAX_GRUPPO_LMC = 5

# Nomi dei fogli del file di input
FOGLIO_STUDENTI = "Studenti"
FOGLIO_DOCENTI = "Docenti"
FOGLIO_GRUPPI = "Gruppi LMC"
FOGLIO_LMI = "LMI"
FOGLIO_PARAMETRI = "Parametri"
FOGLIO_IMPEGNI = "Impegni studenti"  # ore in cui il ragazzo NON può esserci (al contrario dei docenti)
FOGLIO_ABBINAMENTI = "Abbinamenti fissi"  # lezioni già decise: docente, studente, giorno e ora
FOGLI_DATI = [FOGLIO_STUDENTI, FOGLIO_IMPEGNI, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_ABBINAMENTI,
              FOGLIO_LMI, FOGLIO_PARAMETRI]
FOGLI_FACOLTATIVI = {FOGLIO_IMPEGNI, FOGLIO_ABBINAMENTI}  # se mancano si creano vuoti: i file di prima non li hanno

# Come si chiamano i tipi di lezione nel foglio degli abbinamenti
NOMI_TIPO = {"1° strumento": "S1", "2° strumento": "S2", "Musica da camera": "LMC",
             "Laboratorio LMI": "LMI"}
# scritture accettate oltre a quelle sopra (tutto maiuscolo, senza spazi doppi)
_ALIAS_TIPO = {"S1": "S1", "1": "S1", "1°": "S1", "1 STRUMENTO": "S1", "PRIMO STRUMENTO": "S1",
               "S2": "S2", "2": "S2", "2°": "S2", "2 STRUMENTO": "S2", "SECONDO STRUMENTO": "S2",
               "LMC": "LMC", "MUSICA DA CAMERA": "LMC",
               "LMI": "LMI", "LABORATORIO": "LMI", "LABORATORIO LMI": "LMI",
               "LABORATORIO D'INSIEME": "LMI"}


def tipo_da_testo(testo: str) -> str:
    """Il tipo di lezione scritto nel foglio → S1 / S2 / LMC / LMI; "" se non si riconosce."""
    t = " ".join((testo or "").replace("’", "'").split()).upper()
    if not t:
        return ""
    for nome, codice in NOMI_TIPO.items():
        if nome.upper() == t:
            return codice
    return _ALIAS_TIPO.get(t, "")
FOGLIO_ORARIO = "Orario calcolato"  # scritto dal programma dopo il calcolo, riusato al ricalcolo
FOGLIO_TRASPORTI = "Trasporti"  # scritto da "Aggiorna trasporti": minuti di ritorno a casa per fascia

# Orari di partenza da scuola usati per interrogare i mezzi pubblici: fine fascia + 5 minuti
ORE_PARTENZA = ["14:35", "15:35", "16:35", "17:35"]
MINUTI_SENZA_MEZZO = 300  # oltre questi minuti (o senza itinerario) la fascia è "senza mezzo per tornare"

# Tipi di lezione nell'orario prodotto
TIPO_STRUM1 = "S1"
TIPO_STRUM2 = "S2"
TIPO_LMC = "LMC"
TIPO_LMI = "LMI"   # laboratorio d'insieme: normalmente al mattino, si colloca solo se fissato a mano
TIPO_ACCOMP = "ACC"

ETICHETTA_ACCOMP = "Pianista accomp."
