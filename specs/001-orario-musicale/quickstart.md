# Quickstart — Validazione Orario Pomeridiano Liceo Musicale

**Feature**: `001-orario-musicale`
**Date**: 2026-06-29

Questa guida descrive come validare manualmente che l'app funzioni correttamente
per le tre user story principali. Ogni scenario è indipendentemente eseguibile.

## Prerequisiti

```bash
pip install customtkinter pandas reportlab ortools
python main.py
```

La cartella `OrarioMusicale/` viene creata automaticamente alla prima esecuzione
nella directory corrente.

## Dati di Test Minimali

Creare la cartella `OrarioMusicale/test-validazione/` e i seguenti tre file:

### studenti.csv

```
cognome;nome;classe;km;docente_1str;strumento_1;docente_2str;strumento_2;ore_consecutive;giorno_unico
Rossi;Mario;3;12;Bianchi;Violino;Verdi;Teoria;NO;NO
Neri;Sara;5;3;Ferrari;Pianoforte;;;NO;NO
Conti;Luca;1;8;Bianchi;Violino;Verdi;Solfeggio;SI;NO
Bruni;Elena;3;15;Ferrari;Pianoforte;Verdi;Teoria;NO;SI
Galli;Marco;4;2;Bianchi;Violino;Ferrari;Pianoforte;NO;NO
```

### gruppi_lmc.csv

```
numero_gruppo;docente;studente_1;studente_2;studente_3;studente_4
1;Verdi;Rossi;Bruni;Galli;
```

### disponibilita_docenti.csv

```
docente;lun_1330;lun_1430;lun_1530;lun_1630;mar_1330;mar_1430;mar_1530;mar_1630;mer_1330;mer_1430;mer_1530;mer_1630;gio_1330;gio_1430;gio_1530;gio_1630;ven_1330;ven_1430;ven_1530;ven_1630
Bianchi;X;X;X;X;X;X;X;X;;;;X;X;X;;X;X;X;X
Ferrari;;X;X;X;;X;X;X;X;X;X;X;;X;X;X;;X;X;X
Verdi;X;X;;X;X;X;;X;X;X;;X;X;;X;X;X;X;;X
```

---

## Scenario 1: Importazione e gestione dati (User Story P2)

**Obiettivo**: Verificare importazione CSV, editing in-app, esportazione e
persistenza.

1. Avviare l'app: `python main.py`
2. Dalla Home, cliccare "Nuovo anno" → inserire "test-validazione" → confermare
3. Navigare a **Schermata Studenti** → cliccare "Importa CSV" → selezionare
   `OrarioMusicale/test-validazione/studenti.csv`
4. ✅ **Verifica**: 5 studenti visibili in tabella con tutti i campi corretti
5. Doppio-click sulla cella `km` di "Rossi" → cambiare da 12 a 10 → Enter
6. ✅ **Verifica**: valore 10 visibile in tabella
7. Cliccare "Esporta CSV" → aprire il file esportato
8. ✅ **Verifica**: il file contiene `km=10` per Rossi
9. Ripetere passi 3–8 per **Schermata Gruppi LMC** e **Disponibilità docenti**
10. Chiudere l'app → riaprire → selezionare "test-validazione"
11. ✅ **Verifica**: tutti i dati modificati persistono dopo il riavvio

---

## Scenario 2: Generazione orario e export (User Story P1)

**Obiettivo**: Verificare che il solver produca un orario valido e che PDF e
CSV siano esportabili.

**Precondizione**: Dati dell'anno "test-validazione" caricati (Scenario 1 completato).

1. Navigare alla schermata **Genera orario**
2. Cliccare "Genera orario"
3. ✅ **Verifica anteprima**: griglia con colonne Bianchi, Ferrari, Verdi e
   20 righe orarie (lun–ven × 4 fasce)

**Vincoli hard da verificare nell'anteprima**:

| Vincolo | Come verificare |
|---------|-----------------|
| No doppia prenotazione | Nessun docente compare due volte nella stessa fascia |
| Disponibilità rispettata | Controllare alcune celle: es. Bianchi non deve avere lezioni in mer_1330..mer_1630 |
| Co-scheduling LMC (gruppo 1) | Rossi, Bruni e Galli hanno tutti una lezione nella **stessa fascia dello stesso giorno** (tipo LMC) |
| Lezioni individuali adiacenti alla LMC | Le lezioni 1str/2str di Rossi, Bruni, Galli sono nello stesso giorno della LMC |
| ore_consecutive=SI (Conti) | Le 2 ore di violino di Conti sono in fasce adiacenti (es. 13:30 e 14:30) nello stesso giorno |
| giorno_unico=SI (Bruni) | Tutte le lezioni di Bruni (1str + 2str + LMC) cadono nello stesso giorno |

4. Cliccare "Esporta PDF" → aprire il file generato
5. ✅ **Verifica**: PDF A4 orizzontale, tabella leggibile con intestazioni
   docenti e righe con etichette fascia oraria (es. "Lunedì 13:30")
6. Cliccare "Esporta CSV" → aprire il file
7. ✅ **Verifica**: prima colonna "Fascia", colonne successive = docenti,
   20 righe dati, contenuto celle nel formato "COGNOME (cl.X)"

---

## Scenario 3: Vincolo hard non soddisfacibile

**Obiettivo**: Verificare il comportamento su infeasibilità.

1. Nella schermata **Disponibilità docenti**, rimuovere tutte le X di Bianchi
   (o importare un CSV con Bianchi senza disponibilità)
2. Navigare a **Genera orario** → cliccare "Genera orario"
3. ✅ **Verifica**: messaggio di errore esplicito che menziona "Bianchi"
   (es. "Docente Bianchi: disponibilità insufficiente per le lezioni assegnate")
4. ✅ **Verifica**: nessun file PDF o CSV viene creato; i pulsanti Export
   restano disabilitati

---

## Scenario 4: Duplica anno con promozione (User Story P3)

**Obiettivo**: Verificare la promozione automatica.

**Precondizione**: Anno "test-validazione" con dati caricati.

**Studenti attesi dopo la promozione**:

| Cognome | Classe attuale | Classe dopo | Esito |
|---------|---------------|-------------|-------|
| Rossi | 3 | 4 | Promosso |
| Neri | 5 | — | Rimosso (diplomata) |
| Conti | 1 | 2 | Promosso |
| Bruni | 3 | 4 | Promosso |
| Galli | 4 | 5 | Promosso |

1. Dalla Home, selezionare "test-validazione" → cliccare "Duplica anno"
2. Inserire etichetta "test-validazione-2" → cliccare "Avanti"
3. ✅ **Verifica riepilogo**: il wizard mostra la tabella sopra (promossi, rimossi,
   classe 1 vuota)
4. Cliccare "Conferma"
5. Aprire "test-validazione-2" dalla Home
6. ✅ **Verifica studenti**: Rossi cl.4, Conti cl.2, Bruni cl.4, Galli cl.5;
   Neri assente; nessuno in classe 1
7. ✅ **Verifica gruppi LMC**: tabella vuota
8. ✅ **Verifica disponibilità docenti**: identica a "test-validazione"
   (Bianchi, Ferrari, Verdi con le stesse disponibilità)

---

## Scenario 5: Protezione sovrascrittura anno esistente

1. Con "test-validazione-2" già esistente, tentare di duplicare
   "test-validazione" con lo stesso label "test-validazione-2"
2. ✅ **Verifica**: l'app mostra un avviso "Anno test-validazione-2 già
   esistente. Sovrascrivere?" con pulsanti Conferma / Annulla
3. Cliccare "Annulla"
4. ✅ **Verifica**: "test-validazione-2" rimane invariato
