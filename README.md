# Orario Musicale

Programma con finestra per costruire l'orario pomeridiano di un liceo musicale. Si parte da un file
Excel con studenti, docenti e gruppi, e si ottiene l'orario in Excel e PDF.

Funziona senza internet, tranne per la funzione che legge gli orari dei mezzi pubblici.

## Come si usa

```
Apri file Excel  →  correggi nelle griglie  →  Calcola orario  →  Excel + 3 PDF
```

1. Si apre il programma e si sceglie **Apri file Excel** oppure **Nuovo file**.
2. I fogli del file appaiono come griglie modificabili: Studenti, Impegni studenti, Docenti, Gruppi LMC,
   Abbinamenti fissi, LMI, Parametri. Se il file è stato creato con una versione precedente, il programma
   propone di aggiornarlo alla struttura nuova (fogli e colonne mancanti), dopo averne fatto una copia.
3. **Calcola orario** chiede dove salvare e crea una cartella con la data, contenente:
   - `orario.xlsx` — un foglio per giorno (colonne = docenti con aula, righe = le 4 fasce), più i fogli
     Studenti, Docenti, Controlli e LMI
   - `orario_settimanale.pdf` — la griglia da stampare e affiggere, A4 orizzontale
   - `orario_docenti.pdf` — una pagina per docente
   - `orario_studenti.pdf` — le lezioni di ogni ragazzo, per le famiglie
   - gli stessi tre documenti anche in Word (`.docx`), per ritoccarli a mano prima di stamparli
4. Se i dati non tornano, il programma si ferma e mostra l'elenco dei problemi, senza scrivere nulla.

Al calcolo successivo la casella **Parti dall'orario già calcolato** mantiene l'orario e sposta solo le
lezioni che le modifiche rendono necessario spostare.

## Le regole

**Ore settimanali per classe**

| Classe | 1° strumento | 2° strumento | Musica da camera |
|--------|--------------|--------------|------------------|
| 1ª, 2ª | 2 ore | 1 ora | – |
| 3ª, 4ª | 1 ora | 1 ora | 1 ora |
| 5ª | 2 ore | – | 1 ora |

**Fasce**: 13:30, 14:30, 15:30, 16:30, da lunedì a venerdì.

Nei gruppi di musica da camera e nei laboratori i ragazzi si scrivono come **Cognome Nome**, con il menu a
tendina che li suggerisce e il nome aggiunto da solo quando il cognome è di uno solo.

Nel foglio Docenti il pulsante **Compila dagli studenti** aggiunge da solo i docenti che compaiono negli
elenchi degli studenti e nei gruppi, con i loro strumenti. Ogni docente ha **un'aula per giorno**, quindi può
cambiare stanza durante la settimana. La colonna **Ore dichiarate** e la riga **TOTALE** in fondo contano da
sole le ore segnate con la X.

Nel foglio **Impegni studenti** la X vuol dire il contrario che nei docenti: l'ora in cui il ragazzo **non**
può esserci (sport, catechismo…). Il pulsante **Copia i nomi dagli studenti** riempie l'elenco.

Nel foglio **Abbinamenti fissi** si scrivono le lezioni già decise (docente, studente, tipo, giorno, ora): il
programma le blocca e costruisce il resto dell'orario intorno. Al posto dello studente si può scrivere
**«Gruppo 5»** per fissare l'ora di un gruppo di musica da camera, o il nome di un **laboratorio del foglio
LMI**: in quell'ora restano occupati il docente e tutti i ragazzi del gruppo o del laboratorio. Il docente si
può lasciare vuoto, lo prende dal foglio Gruppi LMC o LMI.

Le **2 ore di 1° strumento** (classi 1ª, 2ª e 5ª) vanno di norma in **due giorni diversi**. Per averle una di
seguito all'altra si scrive **SI** nella colonna «1° strumento attaccato» del foglio Studenti; NO o casella
vuota le tiene separate. Fanno eccezione le ore fissate a mano negli abbinamenti e chi ha «Giorno unico = SI».

La colonna **«Minuti per tornare a casa»** la scrive il programma con «Aggiorna trasporti»: è il tempo
migliore fra le quattro fasce, il dettaglio ora per ora resta nel foglio Trasporti.

Chi abita **lontano** salta per quanto possibile l'ultima ora e concentra le lezioni in un pomeriggio solo,
in proporzione al viaggio: quello che conta è **a che ora arriva a casa**, non quanto dura il viaggio. Oltre
la soglia del parametro «Soglia minuti 'tutto in un giorno'» (90 minuti) fa tutto lo stesso giorno, e per lui
la regola delle 2 ore di 1° strumento in giorni diversi non vale.

Su un file compilato con una versione precedente, **Modifica → Riallinea il formato del file** rimette le
istruzioni aggiornate, le larghezze delle colonne e i menu a tendina, senza toccare i dati.

**Non derogabile**: lezioni solo dove il docente è disponibile; un docente e un ragazzo mai in due posti
nella stessa ora; tutti i membri di un gruppo di musica da camera nella stessa ora; giorni vietati per
ragazzo; ore consecutive e giorno unico dove richiesti; ore di pianista accompagnatore in coda alle lezioni
del docente.

**Da ottimizzare, in ordine**: nessun buco tra le lezioni dello stesso pomeriggio, per ragazzi e docenti;
al massimo due rientri a settimana; prime fasce piene e chi abita più lontano collocato prima.

## Trasporti pubblici

Il pulsante **🚌 Trasporti** (o menu Orario → Aggiorna trasporti) calcola, per ogni ragazzo, quanto tempo
serve per tornare a casa se la lezione finisce alle 14:30, 15:30, 16:30 o 17:30. Usa
[Photon](https://photon.komoot.io) e [Nominatim](https://nominatim.openstreetmap.org) per gli indirizzi e
[Transitous](https://transitous.org) per bus e treni: servizi gratuiti e open source, senza chiavi né
registrazioni. I risultati finiscono nel foglio Trasporti del file, quindi il calcolo dell'orario resta
offline e si riusa l'anno dopo.

Il criterio è che più il viaggio di ritorno è lungo, prima il ragazzo va messo in orario. Le fasce dopo le
quali non c'è un mezzo per tornare a casa vengono evitate. Chi non ha l'indirizzo resta ai chilometri.

## Installazione

**macOS**

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

Per l'applicazione da doppio clic: `./costruisci_app.sh` → `dist/Orario Musicale.app`.

**Windows**: per costruire il programma vedere [WINDOWS.md](WINDOWS.md); per installarlo e usarlo,
[GUIDA INSTALLAZIONE WINDOWS.md](GUIDA%20INSTALLAZIONE%20WINDOWS.md). In breve, `costruisci_installer.bat` produce un `setup.exe`
da consegnare, `windows_avvia.bat` prepara l'ambiente e avvia il programma da sorgente.

## Provarlo

Il file `esempio/orario_2026-27.xlsx` contiene un dataset **inventato** già pronto: 30 ragazzi, 7 docenti,
6 gruppi di musica da camera, due laboratori e qualche eccezione. Si apre e si calcola in pochi secondi.
Lo rigenera `dati_esempio.py`.

## Com'è fatto

Python 3.13. Interfaccia con [customtkinter](https://customtkinter.tomschimansky.com) e
[tksheet](https://github.com/ragardner/tksheet), Excel con [openpyxl](https://openpyxl.readthedocs.io),
PDF con [ReportLab](https://www.reportlab.com), calcolo con il risolutore di vincoli
[OR-Tools CP-SAT](https://developers.google.com/optimization/cp/cp_solver).

```
app.py                  avvio
orario/costanti.py      giorni, fasce, ore per classe
orario/modello.py       strutture dati
orario/lettura.py       Excel ↔ dati; fogli "Orario calcolato" e "Trasporti"
orario/template.py      creazione del file di input
orario/controlli.py     controlli preliminari, con messaggi in italiano
orario/motore.py        risolutore CP-SAT
orario/trasporti.py     tempi di ritorno a casa con i mezzi
orario/export_*.py      orario.xlsx e i tre PDF
orario/gui.py           finestra
```

Contesto, decisioni prese e stato del lavoro sono in [PROGETTO_ORARIO.md](PROGETTO_ORARIO.md).

## Dati personali

I file con nomi, indirizzi e recapiti dei ragazzi **non stanno in questa repository**: sono esclusi dal
`.gitignore` e restano solo sul computer di chi costruisce l'orario. I dati di esempio sono inventati.
