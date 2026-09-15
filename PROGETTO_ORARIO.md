# Orario pomeridiano Liceo Musicale – Documento di progetto

Ultimo aggiornamento: 8 settembre 2026 (sera): prima versione funzionante del programma.
Questo file raccoglie tutto quanto discusso con Francesco: contesto, dati, decisioni prese, domande aperte
e piano di lavoro. Va tenuto aggiornato a ogni sessione di progettazione.

---

## 1. Obiettivo

Un **programma con finestra, semplice**, senza gestione di più anni e senza database. Il flusso è uno solo:

1. Si apre il programma. Si sceglie **Apri file Excel…** oppure **Nuovo file…** (crea un file di input vuoto,
   con righe grigie di esempio da cancellare, e lo apre).
2. Il programma lo importa e lo **mostra a video come un foglio di calcolo**: una scheda per foglio
   (Studenti, Docenti, Gruppi LMC, LMI, Parametri), celle modificabili per correzioni al volo.
3. Si preme **Calcola orario**. Il programma chiede **dove salvare i risultati** e lì crea una cartella
   "Risultati orario AAAA-MM-GG HH.MM" con Excel e PDF. Prima controlla i dati; se qualcosa non torna, si
   ferma e mostra l'elenco dei problemi (vedi sez. 6).
4. In alto c'è la barra dei menu: **File** (Nuovo, Apri, Riapri ultimo, Salva, Salva con nome, Chiudi, Esci),
   **Modifica** (aggiungi/elimina righe), **Orario** (Calcola, Ricalcola da zero, Apri ultima cartella dei
   risultati), **Aiuto** (Istruzioni, Informazioni).

Le modifiche fatte a video si possono salvare nel file Excel di input (pulsante **Salva**), così il file
resta la fonte unica e si può riaprire l'anno dopo.

**Ricalcolo con modifiche minime** (richiesta dell'8/9/2026): dopo ogni calcolo il programma scrive l'orario
anche in un foglio del file di input, "Orario calcolato". Se si cambia qualcosa (disponibilità di un docente,
uno studente in più…) e si ricalcola, il programma parte da quell'orario e **sposta solo le lezioni
strettamente necessarie**, tenendo coerente il resto. Una casella "Parti dall'orario già calcolato" permette
di disattivarlo e ricalcolare da zero. Gli spostamenti vengono elencati negli avvisi.

La versione precedente del progetto (app a 6 schermate con gestione anni e wizard "duplica anno",
specifica in `specs/001-orario-musicale/`) è **superata** da questo documento. Le tabelle a griglia già
scritte in `gui/studenti.py`, `gui/gruppi_lmc.py`, `gui/disponibilita.py` possono essere riutilizzate;
`gui/home.py` (lista anni) e la logica multi-anno di `main.py` vanno tolte.

## 2. Il problema da risolvere (regole della scuola)

**Fasce orarie.** 4 rientri al giorno, da lunedì a venerdì, quindi 20 fasce a settimana:

| Ora | Fascia |
|-----|--------|
| 7ª  | 13:30 – 14:30 |
| 8ª  | 14:30 – 15:30 |
| 9ª  | 15:30 – 16:30 |
| 10ª | 16:30 – 17:30 |

**Ore settimanali per ragazzo, in base alla classe:**

| Classe | 1° strumento | 2° strumento | Musica da camera (LMC) |
|--------|--------------|--------------|------------------------|
| 1ª, 2ª | 2 ore | 1 ora | – |
| 3ª, 4ª | 1 ora | 1 ora | 1 ora |
| 5ª     | 2 ore | – | 1 ora |

**Gruppi LMC**: da 2 a 5 ragazzi delle classi 3ª, 4ª, 5ª (anche misti), con un docente. Tutti i membri
del gruppo devono avere la lezione nella stessa fascia dello stesso giorno.

**Glossario.** *LMC* = Laboratorio di Musica da Camera: 1 ora a settimana, piccoli gruppi (2–5) del
triennio, **nel pomeriggio**: è oggetto di questo programma. *LMI* = Laboratorio di Musica d'Insieme: 2 ore a
settimana per tutte le classi, grandi gruppi per famiglia di strumento (archi, chitarre, fiati, canto,
percussioni), **nell'orario del mattino** (es. 1ª: martedì 4ª e 5ª ora). Gli LMI **non entrano** nell'orario
pomeridiano e il programma li ignora; contano solo nel monte ore dei docenti (es. "Bini 16 ore + 2 LMI"),
che però non è un vincolo del programma: le disponibilità pomeridiane le danno i docenti con le X.

**Docenti**: ogni docente comunica le fasce in cui è disponibile. Un docente non può avere due lezioni
nella stessa fascia. Le lezioni di strumento e quelle di LMC dello stesso docente usano **la stessa
disponibilità** (deciso l'8/9/2026).

**Ragazzi**:
- Pochissimi buchi: le lezioni dello stesso pomeriggio devono essere una di seguito all'altra.
- Al massimo 2 rientri a settimana; possono diventare 3 per chi abita vicino alla scuola.
- Chi abita lontano va nelle prime fasce, chi abita vicino nelle ultime (dipende dai mezzi pubblici).
- Casi particolari: qualcuno deve rientrare un solo giorno; qualcuno ha le 2 ore di 1° strumento consecutive.

**Come lavora Francesco a mano oggi**: colloca prima i gruppi LMC, poi mette le lezioni individuali dei
ragazzi del gruppo nelle ore subito vicine. Con il programma questo passaggio non serve: basta chiedere
di minimizzare i buchi e i rientri, e il risultato viene da solo.

---

## 3. I file di partenza e cosa contengono

> **Nota (8/9/2026)**: i tre file qui sotto sono **dati di prova**, utili per capire il problema e provare il
> programma. Il dataset reale 2026-27 arriva il 9/9/2026. Le mancanze elencate in 3.4 (docente Percussioni 2,
> gruppi LMC, disponibilità, km della 1ª) si risolvono con i dati veri, non con domande.

### 3.1 `Recapiti musicale 2025.26 con distanze.xls`
Un foglio, 102 ragazzi. Colonne: Cognome, Nome, Classe, Comune, Provincia, Indirizzo, Stato alunno, KM.
La colonna **KM** (distanza casa-scuola) è l'unica che serve al programma. È l'elenco dell'**anno scorso**:
le classi indicate sono quelle del 2025-26 e mancano i nuovi iscritti in 1ª.

### 3.2 `classi e gruppi 2026-27 AGGIORNATO AL 07-4.docx`
È il file dell'**anno nuovo**. Tre blocchi:

1. **Le cinque classi**, una tabella per classe: cognome, nome, 1° strumento e docente, 2° strumento e
   docente (la 5ª ha solo il 1° strumento). Totale 109 ragazzi: 1ª=28, 2ª=13, 3ª=26, 4ª=22, 5ª=20.
   **Nella 1ª i docenti sono tutti vuoti** ("completamenti secondo disponibilità").
2. **Tabelle per strumento** (pianoforte, fiati, chitarra, archi, canto): riepilogo di quanti ragazzi ha
   ogni docente e quante ore. Stessa informazione delle classi, vista dal lato docente.
3. **Griglia LMC** (ultima pagina): 25 gruppi con nomi e docente. **È quella dell'anno scorso**: contiene
   ragazzi ormai diplomati (Gianassi, Zaho, Reali, Porrino, Azzurri, Ciofi…). Per il 2026-27 va rifatta.

Particolarità nei nomi dei docenti: "Pacini C" (pianoforte) e "Pacini M" (chitarra) sono due persone
diverse; "Perc 2" (Bardi, Petrone, Sali, Magherini, Sestito) non è un cognome ma la cattedra "Percussioni 2",
il cui docente non è ancora nominato ("Percussioni 1" è Mazzei). Nel foglio Docenti resta come segnaposto da
sostituire con il nome quando si saprà. L'oboe è Francini.

### 3.3 `ORARIO POMERIDIANO DEFINITIVO 2025-26.pdf`
L'orario finito dell'anno scorso, cioè **il modello dell'output**. Una pagina per giorno. Una colonna
per docente con in alto il nome dell'aula (1ALM, M0, M1…). Quattro righe, una per fascia (7ª–10ª). In ogni
casella: cognome del ragazzo + classe; grassetto = 1° strumento, corsivo = 2° strumento. La lezione di
gruppo è una casella "LMC" con tutti i cognomi. Alcune caselle: "Lezione di 2 ore", "Pianista accomp." (solo
nella colonna di Paccosi, 5 ore a settimana: il docente di pianoforte accompagna al piano le lezioni o le prove
di altri ragazzi, cantanti e strumentisti, e in quell'ora non fa lezione individuale; nel Word è il "+ 1
accompagnamento" accanto al suo monte ore).
Ogni docente ha un'aula fissa per tutta la settimana.

### 3.4 Cosa manca
- **Disponibilità dei docenti**: nessun file. Si compila nel foglio "Docenti" di `input_orario.xlsx`.
- **Gruppi LMC 2026-27**: da comporre a mano nel foglio "Gruppi LMC".
- **Docenti della 1ª**: da assegnare a mano nel foglio "Studenti".
- **KM dei 28 ragazzi di 1ª**: da aggiungere quando arrivano i recapiti nuovi.

---

## 4. Il file di input: `input_orario.xlsx`

Si crea dal programma con **Nuovo file**, oppure si parte da quello dell'anno precedente.
Si può rilanciare quando il Word cambia (attenzione: sovrascrive il file, quindi le compilazioni a mano
vanno rifatte o copiate; in futuro si può fare uno script di aggiornamento che preserva le modifiche).

| Foglio | Contenuto | Stato precompilazione |
|--------|-----------|-----------------------|
| Istruzioni | Come compilare | completo |
| Studenti | Classe, Cognome, Nome, Strumento 1, Docente 1, Strumento 2, Docente 2, 1° strumento attaccato (SI/NO), Giorno unico (SI/NO), Giorni NON disponibili, Comune, Indirizzo, Civico, KM, Minuti per tornare a casa, Note | 109 ragazzi; 81 con KM; le celle da completare sono **gialle** |
| Impegni studenti | Classe, Cognome, Nome + 20 colonne Lun 13:30 … Ven 16:30 (X = il ragazzo **non** può esserci), Note | vuoto; il pulsante «Copia i nomi dagli studenti» lo riempie |
| Docenti | Docente, Strumento/i, **5 aule** (una per giorno), **Ore accompagnamento** (numero), 20 colonne Lun 13:30 … Ven 16:30 (X = disponibile, vuoto = no, A = ora di accompagnamento fissata a mano), **Ore dichiarate**, Note; in fondo la riga **TOTALE** | 21 docenti; aule dell'anno scorso; disponibilità **vuote** |
| Gruppi LMC | Gruppo, Docente, Studente 1…5, Note | vuoto, con 2 righe di esempio grigie da cancellare; i gruppi li fornisce lo zio |
| Abbinamenti fissi | Docente, Studente (oppure «Gruppo N» di musica da camera, oppure un laboratorio del foglio LMI), Tipo di lezione, Giorno, Ora, Note | vuoto; le lezioni decise a mano, che il motore blocca |
| LMI | Laboratorio, Classi, Docente, Aula, Giorno e ora (mattino), Studenti, Note | vuoto, con 1 riga di esempio; il programma non li calcola, li ricopia in output |
| Parametri | Max rientri (2), max rientri chi abita vicino (3), soglia KM "vicino" (5), tempo massimo di calcolo (120 s), indirizzo della scuola, soglia minuti "vicino" (25) | valori proposti; indirizzo scuola da compilare |
| Trasporti | scritto dal programma (Aggiorna trasporti): per studente, minuti/arrivo/mezzi per fascia | vuoto finché non si aggiorna |

Le colonne Docente e Studente hanno menu a tendina collegati agli altri fogli, così i nomi coincidono.
I fogli aggiunti dopo (Impegni studenti, Abbinamenti fissi) mancano nei file più vecchi: all'apertura il
programma se ne accorge e propone di aggiornarli, vedi «Aggiornamento dei file delle versioni precedenti».

---

## 5. I file di output (da realizzare)

Salvati nella cartella scelta al momento del calcolo, dentro una sottocartella "Risultati orario AAAA-MM-GG HH.MM":

| File | Contenuto |
|------|-----------|
| `orario.xlsx` | Fogli **Lun, Mar, Mer, Gio, Ven**: stesso impianto del PDF 2025-26 (colonne = docenti con aula, righe = 4 fasce, celle = cognome + classe, grassetto/corsivo per 1°/2° strumento, celle LMC con tutti i cognomi, lezioni di 2 ore evidenziate). Foglio **Studenti**: un rigo per ragazzo con i suoi rientri e le sue ore. Foglio **Docenti**: la settimana di ciascun docente. Foglio **Controlli**: rientri, buchi, eccezioni alla regola della distanza, tutto ciò che il programma ha dovuto forzare. Foglio **LMI**: i laboratori del mattino ricopiati dal foglio di input, senza calcolo. |
| `orario_settimanale.pdf` | La griglia per giorno, A4 orizzontale, una pagina per giorno, pronta da stampare e affiggere (come il PDF 2025-26). In coda una pagina con i **LMI del mattino** (docente, aula, giorno e ora, ragazzi), come il riquadro in fondo al martedì del PDF vecchio. |
| `orario_docenti.pdf` | Una pagina per docente con la sua settimana. |
| `orario_studenti.pdf` | Una pagina (o mezza) per ragazzo, da consegnare alle famiglie. |
| `orario_settimanale.docx`, `orario_docenti.docx`, `orario_studenti.docx` | Gli stessi tre documenti in Word, con la stessa impaginazione, per ritoccarli a mano prima di stamparli o inviarli (richiesta del 14/9/2026). |

Dopo il calcolo il programma mostra a video l'anteprima della griglia settimanale e apre la cartella dei file.

## 6. Come funziona il motore (spiegazione semplice)

Si usa un **risolutore di vincoli** (OR-Tools CP-SAT, libreria gratuita di Google). Gli si descrivono:

**Regole rigide** (mai violate):
1. Una lezione solo nelle fasce in cui il docente ha la X o la A.
1b. Ogni docente ha esattamente il numero indicato di ore di accompagnamento, collocate in fasce con X
    (o A, che le fissa) e mai sovrapposte a una sua lezione.
1c. Le ore di accompagnamento stanno sempre **in coda** alle ore del docente nella giornata: prima le lezioni,
    poi l'accompagnamento (deciso l'8/9/2026).
2. Un docente non ha mai due lezioni nella stessa fascia.
3. Un ragazzo non ha mai due lezioni nella stessa fascia.
4. Ogni ragazzo riceve esattamente le ore previste dalla sua classe.
5. Tutti i membri di un gruppo LMC nella stessa fascia dello stesso giorno, con il loro docente.
6. Rispetto dei giorni in cui il ragazzo non può venire.
7. 1° strumento attaccato = SI → le 2 ore di 1° strumento sono adiacenti nello stesso giorno; = NO → in due giorni diversi.
8. Giorno unico = SI → tutte le lezioni del ragazzo nello stesso giorno.

**Preferenze** (il programma cerca di soddisfarle il più possibile, in quest'ordine di importanza):
1. Zero buchi tra le lezioni dello stesso pomeriggio.
2. Massimo 2 rientri (3 per chi abita vicino).
3. **Prime fasce piene, distanza progressiva** (deciso l'8/9/2026, corretto la sera dopo la prima prova):
   per tutti costa un po' stare tardi, così si riempiono prima le 13:30 e le 14:30 e le ultime fasce restano
   vuote dove possibile. Per chi abita lontano il costo cresce con i KM, senza soglie fisse: il programma
   "parte dai più lontani", e i vicini finiscono nelle fasce tardive solo quando non c'è più posto prima,
   non per scelta.

4. **Ore individuali attaccate alla musica da camera** (richiesta del 14/9/2026): per i ragazzi del
   triennio le ore di strumento vanno nello stesso pomeriggio della lezione di gruppo, subito prima o subito
   dopo. Non essendoci buchi, "stesso giorno" equivale ad "attaccate". Peso 700, sotto quello di un buco:
   il programma non crea un vuoto pur di attaccarle. Sui dati di esempio si passa dal 25% al **91%** di ore
   attaccate (29 su 32), sempre con zero buchi, e i ragazzi che vengono un solo pomeriggio passano da 2 a 13.
   Le poche ore che non si riesce ad attaccare finiscono negli avvisi.
5. **Buchi dei docenti**: anche la giornata del docente deve essere compatta (obiettivo generale del
   programma: meno buchi possibile per tutti, ragazzi e docenti). Le ore di accompagnamento si attaccano
   alle lezioni, senza spazi vuoti.

**Se qualcosa non torna** (deciso l'8/9/2026): il programma **si ferma e non produce l'orario**. Elenca
invece i problemi, nel modo più leggibile possibile, per esempio:
- "Rossi (3ª): il docente Bini e il docente Corsini non hanno nessuna fascia utile in comune con i vincoli
  del ragazzo".
- "Docente Simonelli: 27 ore richieste ma solo 24 fasce disponibili".
- "Gruppo LMC 4: lo studente 'Manzi' non esiste nel foglio Studenti".
Prima di far partire il calcolo vero, il programma fa una serie di controlli preliminari (nomi che non
combaciano, ore richieste > disponibilità, ecc.) e li segnala tutti insieme.

---

## 7. Decisioni prese

| Data | Decisione |
|------|-----------|
| 8/9/2026 | Versione semplice: programma con finestra a un solo flusso (apri file → correggi a video → Calcola → Excel + PDF). Nessuna gestione multi-anno, nessun database. |
| 8/9/2026 | Un solo file di input con più fogli (Studenti, Docenti con disponibilità, Gruppi LMC, Parametri), precompilato dai file esistenti. |
| 8/9/2026 | Distanza gestita in modo progressivo per KM, senza soglie fisse; si parte dai più lontani. |
| 8/9/2026 | In caso di problemi il programma si ferma e li elenca; niente orari parziali. |
| 8/9/2026 | Lezioni di strumento e LMC dello stesso docente condividono la stessa disponibilità. |
| 8/9/2026 | I gruppi LMC vanno da 2 a 5 ragazzi (la griglia vecchia arrivava a 4). |
| 8/9/2026 | Disponibilità docenti: griglia docente × 20 fasce con X (foglio Docenti). Confermato. |
| 8/9/2026 | "Perc 2" = cattedra Percussioni 2, docente da nominare; resta segnaposto nel foglio Docenti. |
| 8/9/2026 | Nell'orario si scrive cognome + classe; iniziale del nome solo per cognomi uguali nella stessa classe. Confermato. |
| 8/9/2026 | Struttura del file `input_orario.xlsx` approvata da Francesco. |
| 8/9/2026 | I gruppi LMC 2026-27 li fornisce lo zio (coordinatore); il programma li colloca soltanto. |
| 8/9/2026 | Gli LMI (mattino) entrano nell'output come pagina ricopiata dal foglio LMI di input, senza calcolo. |
| 8/9/2026 | Aule: ogni docente ha la sua aula fissa per tutta la settimana; si scrive in testa alla colonna, nessun controllo di conflitto. |
| 8/9/2026 | I dati attuali sono di prova; il dataset reale arriva il 9/9/2026. Niente domande allo zio per chiedere dati. |
| 8/9/2026 | **Assunzione** (da confermare con lo zio): i docenti della 1ª li assegna lui a mano prima del calcolo. |
| 8/9/2026 | Pianista accompagnatore: per ogni docente si indica il **numero di ore** di accompagnamento a settimana; il programma le colloca nelle ore libere del docente, **sempre in coda alle sue lezioni della giornata**, e stampa "Pianista accomp." nella casella. Una **A** nella griglia fissa a mano una di quelle ore (doppia modalità confermata da Francesco). |
| 8/9/2026 | Obiettivo generale: meno buchi possibile per tutti, ragazzi e docenti. |
| 8/9/2026 | Ricalcolo stabile: l'orario calcolato si salva nel foglio "Orario calcolato" del file di input; al ricalcolo il programma sposta solo il necessario. |
| 8/9/2026 | Il calcolo usa tutti i core della CPU (anche su Windows). Menu in alto, "Nuovo file", scelta della cartella dei risultati con data (richieste di Francesco dopo la prima prova). |
| 8/9/2026 | Stack: Python 3.13, customtkinter (finestra), tksheet (griglie tipo Excel), openpyxl, OR-Tools CP-SAT, reportlab, PyInstaller. Codice nel pacchetto `orario/`, avvio con `app.py`. |

---

## 8. Domande aperte allo zio coordinatore (messaggio inviato l'8/9/2026)

Lo zio si chiama Giovanni. Si procede con le assunzioni in tabella 7; le domande servono a confermarle.

1. **Docenti della 1ª**: li assegna lui a mano, oppure vuole che li distribuisca il programma?
   Si procede con: *a mano*.
2. **Pianista accompagnatore**: si procede con il numero di ore per docente, distribuite dal programma
   nelle ore libere (più eventuali A fissate a mano). Chiesto allo zio: le ore di accompagnamento devono
   coincidere con la lezione di un ragazzo preciso (es. il cantante accompagnato)? Se sì, si aggiunge una
   colonna "accompagna lo studente" e il programma le fa coincidere.
3. **LMI**: vuole la pagina dei laboratori del mattino in coda all'orario stampato, come l'anno scorso?
   Si procede con: *sì* (foglio LMI in input, pagina LMI in output, solo ricopiata).

---

## 8b. Mezzi pubblici al posto dei KM (richiesta dell'8/9/2026, realizzato la sera stessa)

Obiettivo: ridurre la permanenza a scuola dei ragazzi che vengono da fuori, in base agli orari reali dei
mezzi (bus Autolinee Toscane, treni regionali da Pistoia), non solo ai chilometri.

**Idea di fondo**: ciò che conta per il programma non è l'indirizzo ma, per ogni ragazzo, "quanto costa"
finire la lezione alle 14:30, 15:30, 16:30 o 17:30 (attesa del mezzo + viaggio) e quali fasce sono
impossibili (ultimo mezzo). Il motore già lavora con un costo per fascia per studente: basta cambiare
la sorgente del costo.

**Fase A – tabella "Trasporti" per comune, compilata dal coordinatore** (fattibile subito, offline):
un foglio con un rigo per comune/frazione (~15 righe): orari di partenza dei mezzi da Pistoia nel
pomeriggio (es. "14:40, 17:40, 19:10"), minuti di viaggio, eventuale ultima partenza utile. Il programma
calcola per ogni ragazzo l'attesa dopo ogni fascia e la trasforma in costo; le fasce oltre l'ultimo mezzo
diventano vietate. I ragazzi di Pistoia restano "vicini" (nessun costo di attesa). I KM restano come
ripiego per chi non ha il comune in tabella.

**Fase B – orari automatici dagli indirizzi** (richiede internet e un account): un pulsante "Aggiorna
trasporti" che interroga un servizio di percorsi con mezzi pubblici (Google Maps Directions in modalità
transit è il più affidabile per la Toscana; costo trascurabile per ~100 ragazzi, ma serve un account
Google Cloud con carta) e salva i risultati nel foglio Trasporti. Il calcolo dell'orario resta offline.
Alternativa gratuita: dati GTFS aperti di Autolinee Toscane e Trenitalia con un motore di percorsi
locale, molto più lavoro e fragile. **Privacy**: inviare al servizio solo comune/frazione (o la fermata),
non l'indirizzo esatto dei minori: la precisione a livello di paese basta.

**Decisione (8/9/2026, sera)**: si fa la Fase B con servizi **open source e gratuiti, senza chiave né carta**:
- **Nominatim (OpenStreetMap)** per trasformare gli indirizzi in coordinate (1 richiesta al secondo, con
  User-Agent identificativo). Provato sugli indirizzi veri: Larciano, Cerreto Guidi, Cantagallo,
  Lamporecchio, Pistoia, San Marcello trovati; il solo nome del comune montano dà un punto inutile, serve
  la via.
- **Transitous** (api.transitous.org, motore MOTIS, comunità open source) per i percorsi con i mezzi:
  conosce Autolinee Toscane (urbano ed extraurbano di Pistoia, Lucca, Firenze) e Trenitalia. Provato:
  Pistoia→Sambuca bus 56 (60 min), →Pescia treno+bus (53–81 min), →Cerreto Guidi 125 min,
  →Cantagallo 86 min, →Larciano 70 min, →San Marcello bus 54 (52 min), →Pistoia Pontelungo 10 min.
- Criterio: **più il viaggio di ritorno è lungo, più il ragazzo va messo presto** (stesso criterio
  progressivo dei km, ma sui minuti di viaggio + attesa reali per ciascuna fascia). Le fasce dopo
  l'ultimo mezzo diventano vietate. I km restano come ripiego se il servizio non trova nulla.
- Rischi accettati: servizio comunitario senza garanzie (se sparisce, lo stesso motore MOTIS si può
  installare in casa con i dati GTFS aperti di Autolinee Toscane); i risultati si salvano nel file, quindi
  il calcolo dell'orario resta offline. Privacy: dati trattati solo sul computer del docente autorizzato e
  mai stampati (valutazione di Francesco).

**Come funziona nel programma (realizzato)**
- Foglio Studenti: colonne **Comune, Indirizzo, Civico** (da compilare con comune, via e civico).
  KM resta come ripiego.
- Foglio Parametri: **Indirizzo della scuola** (punto di partenza) e **Soglia minuti "abita vicino"** (25).
- Pulsante **🚌 Trasporti** nella barra in alto (e voce **Orario → Aggiorna trasporti** nel menu). Il pulsante
  mostra la copertura: "60/80" in blu se mancano dei ragazzi, "✓ 80" in verde quando sono tutti a posto,
  disattivato se nessuno ha l'indirizzo. Inoltre, premendo **Calcola orario**, se ci sono indirizzi senza tempi
  dei mezzi il programma lo dice e offre due strade: calcolare prima i trasporti mancanti e poi proseguire da
  solo con l'orario, oppure calcolare subito usando i KM. Funzionamento: per ogni studente con indirizzo, geocodifica prima con **Photon**
  (komoot, dati OpenStreetMap: veloce, senza limite di 1 richiesta/secondo, capisce le abbreviazioni) e con
  **Nominatim** solo come riserva (indirizzo con civico → via → solo nome della frazione → comune; accettato
  solo se nel comune giusto, senza farsi ingannare dalla provincia "Pistoia"; trattini e apostrofi nei nomi
  dei comuni ignorati, es. "Montecatini-Terme"). Poi 4 richieste a Transitous (partenza da scuola alle
  14:35, 15:35, 16:35, 17:35 del prossimo martedì), due studenti alla volta. Salva nel foglio **Trasporti**:
  minuti da fine lezione ad arrivo a casa, ora di arrivo e mezzi, per ogni fascia. Si può interrompere e i
  risultati parziali restano; le coordinate già note si riusano.
  Misure sugli 80 studenti di prova con indirizzo: prima versione (solo Nominatim, 4 richieste a Transitous
  per studente) 1055 s e 73 trovati; versione finale (Photon, e a Transitous si chiedono **tutte le
  partenze del pomeriggio in una richiesta** invece di una per fascia — una seconda serve solo se le corse
  trovate non arrivano all'ultima fascia; 30 min a piedi dalla fermata, centro del comune come ultimo
  ripiego) **249 s e 80 trovati su 80**, cioè circa 3 s per studente. Transitous limita il ritmo a circa una richiesta ogni 3
  secondi per utente: il tempo dipende dal numero di studenti, non dalla velocità del computer.
  **Riuso**: coordinate e percorsi restano nel foglio Trasporti; al prossimo aggiornamento il programma chiede
  se rifare tutti (orari dei mezzi cambiati) o solo i nuovi/modificati (secondi). Provato: 1 studente nuovo
  su 80 → 6 s.
- Motore: la **lontananza** di ogni studente è il miglior tempo di ritorno a casa tra le 4 fasce (o, senza
  trasporti, 10 + 2,5·km); il criterio principale resta quello di Francesco, **chi è più lontano va messo
  prima**, applicato a tutti allo stesso modo (costo crescente con l'ora, più forte per i lontani). Con i dati
  dei mezzi si aggiunge un costo pari ai minuti di attesa+viaggio dopo la lezione (a parità, la fascia da cui
  si torna prima) e una penalità grande per una fascia dopo la quale non c'è un mezzo (o si arriva oltre 5 ore
  dopo). Prima prova con il solo costo "minuti di ritorno" dava l'effetto opposto (i lontani tardi, perché
  il bus è lo stesso a qualunque ora): corretto la sera dell'8/9. Avvisi: "lezione ... ma dopo quell'ora non
  risulta un mezzo per tornare a casa".
- Controlli: avviso se ci sono indirizzi ma i trasporti non sono stati calcolati (si usano i km).
- Codice: `orario/trasporti.py`; foglio letto/scritto da `lettura.leggi_trasporti/salva_trasporti`.

**Dati che servono dal coordinatore**: indirizzo esatto della scuola (foglio Parametri).

---

## 9. Piano di lavoro

1. ✅ Analisi dei file e documento di progetto (questo file).
2. ✅ File di input `input_orario.xlsx` con i dati veri della scuola (compilato una volta).
3. ⬜ Arriva il dataset vero (9/9/2026): si completa `input_orario.xlsx` in Excel o nel programma.
4. ✅ **Finestra 1 – Apertura**: pulsante "Apri file Excel…" e "Riapri ultimo file".
5. ✅ **Finestra 2 – Dati**: schede Studenti / Docenti / Gruppi LMC / LMI / Parametri a griglia (tksheet),
   celle modificabili, aggiungi/elimina righe, pulsanti "Salva" e "Calcola orario".
6. ✅ Controlli preliminari con messaggi chiari (finestra "problemi da risolvere", testo copiabile).
7. ✅ Motore CP-SAT con regole rigide e preferenze (sez. 6); diagnosi leggibile se impossibile.
8. ✅ `orario.xlsx` e i tre PDF (sez. 5); anteprima a video per giorno + scheda avvisi.
9. ✅ Ricalcolo stabile dal foglio "Orario calcolato".
10. ⬜ Prova con i dati veri, confronto con l'orario 2025-26, messa a punto dei pesi.
11. ✅ Pacchetto eseguibile: `.app` per Mac costruito e provato (`costruisci_app.sh`); per Windows `costruisci_app.bat` da eseguire su un PC Windows.
12. ✅ Pulizia: rimosso il vecchio codice multi-anno (gui/, data/, main.py). La spec 001 resta solo come storia.

**Tecnologie**: Python 3.13; customtkinter (finestra); tksheet (griglie tipo Excel); openpyxl (Excel);
reportlab (PDF); OR-Tools CP-SAT (calcolo); PyInstaller (eseguibile).

**Struttura del codice**
```
app.py                  avvio del programma
orario/costanti.py      giorni, fasce, ore per classe, nomi dei fogli
orario/modello.py       strutture dati (Studente, Docente, GruppoLMC, Lezione, Orario, Problema)
orario/lettura.py       Excel ↔ tabelle ↔ DatiInput; foglio "Orario calcolato"
orario/controlli.py     controlli preliminari, risoluzione cognomi nei gruppi
orario/motore.py        risolutore CP-SAT (regole rigide, preferenze, ricalcolo, diagnosi)
orario/export_*.py      orario.xlsx e i tre PDF (export.py = punto unico)
orario/gui.py           finestra (menu, schermata apertura, griglie, calcolo, risultati)
orario/trasporti.py     Nominatim + Transitous: tempi di ritorno a casa per fascia → foglio Trasporti
orario/template.py      costruzione del file di input (vuoto o precompilato)
dati_prova.py           dati veri "riempiti" → input_prova.xlsx (contiene nomi reali: fuori dal repo)
dati_esempio.py         dataset inventato → esempio/orario_2026-27.xlsx (30 studenti, 7 docenti)
test_motore.py, test_export.py, test_e2e.py, test_gui.py   prove da terminale
installer/installer.iss  script Inno Setup per l'installer Windows
```

**Aule per giorno (9/9/2026, richiesta di Francesco)**: nel foglio Docenti l'aula non è più una sola per
tutta la settimana ma cinque, una per giorno (Aula Lun … Aula Ven). L'orario del giorno, i PDF e l'anteprima
mostrano l'aula di quel giorno; nel foglio riepilogativo per docente compare "varia" se cambia, con l'aula
scritta accanto a ogni giornata. I file creati prima, con la colonna "Aula" unica, vengono aggiornati
all'apertura copiando quell'aula su tutti e cinque i giorni.

**Pulsante «Compila dagli studenti» (9/9/2026)**: nel foglio Docenti (e nel menu Modifica). Raccoglie i nomi
che compaiono in Docente 1 e Docente 2 del foglio Studenti e nei gruppi di musica da camera, aggiunge i
docenti mancanti, completa la colonna degli strumenti e lascia intatte aule, disponibilità, ore di
accompagnamento e note. Alla fine dice quanti ne ha aggiunti, a quanti ha completato gli strumenti e quali
docenti non sono usati da nessuno studente, che di solito vuol dire un nome scritto in due modi diversi.

**Nomi degli studenti nei gruppi e negli LMI (9/9/2026)**: nelle celle si scrive **Cognome Nome**, perché
due ragazzi possono avere lo stesso cognome. Nel foglio Gruppi LMC ogni cella ha il menu a tendina con tutti
i ragazzi (si può anche digitare); scrivendo solo il cognome, il nome viene aggiunto da solo quando non ci
sono omonimi, e negli LMI questo vale per ogni nome dell'elenco separato da virgole. Con gli omonimi (nei
dati veri: Gori Camilla di 3ª e Gori Yvaine di 5ª) la cella resta come scritta, così è chi compila a
scegliere. La verifica dei nomi al momento del calcolo accetta entrambe le forme.

**L'orario salvato porta la versione delle regole (15/9/2026, dall'audit di un output)**: l'orario delle
00:35 era peggiore di quello che il motore produce da zero — 8 buchi invece di 1, gradiente per distanza
debole, 36 scambi vicino/lontano ancora possibili, Fiesoli alle 16:30. Il motivo: era stato calcolato con
«Parti dall'orario già calcolato» acceso (si accende da solo), quindi adattando l'orario vecchio, fatto con
le regole vecchie. Ogni lezione tenuta ferma vale 800 punti e spostarne due per uno scambio 1600, più del
guadagno dello scambio (~960): la struttura vecchia si conservava. Ora il foglio «Orario calcolato» riporta
`Regole vN` (`motore.VERSIONE_REGOLE`, da alzare a ogni cambio di criterio); se la versione è diversa da
quella del programma, la casella resta spenta e lo dice: «è di una versione precedente: meglio ricalcolare da
zero». L'adattamento resta la scelta giusta quando cambiano i dati, non quando cambiano le regole.

**Le due ore di 1° strumento mai in giorni di fila (15/9/2026, richiesta di Francesco)**: non basta che
siano in giorni diversi, fra una lezione e l'altra ci vuole almeno un giorno in mezzo (lunedì-mercoledì,
martedì-venerdì…). Erano venute fuori Bonacchi lunedì-martedì e Pippi due volte lo stesso lunedì. Il vincolo
è **cedevole** ma carissimo (`PESO_GIORNI_VICINI = 4000`, sopra il costo di spostare due lezioni in un
ricalcolo): con la regola rigida l'orario non esisteva proprio, perché a un docente con 9 ore dichiarate e
solo giorni vicini non resta nessuna coppia valida. Chi resta vicino finisce negli avvisi con il motivo. Le
eccezioni sono tre: «1° strumento attaccato = SI», «Giorno unico = SI» e le ore fissate a mano.

**Il pomeriggio unico solo a chi lo chiede (15/9/2026, richiesta di Francesco)**: la scorciatoia «chi abita
oltre 90 minuti fa tutto lo stesso giorno» è stata tolta, insieme al parametro che la governava. Due lezioni
dello stesso strumento troppo vicine non hanno senso nemmeno per chi viene da lontano; il pomeriggio unico
resta solo a chi ha «Giorno unico = SI». Per gli altri valgono le preferenze normali, che già concentrano i
rientri di chi viaggia molto.

**Le prime ore a chi viene da lontano, le ultime a chi sta vicino (15/9/2026, richiesta di Francesco)**:
Degl'Innocenti (49 minuti) aveva tre lezioni alla prima ora e Chen (71 minuti) tutte alle ultime. Il criterio
c'era ma era troppo debole: scambiare un vicino e un lontano fra la prima e l'ultima fascia cambiava il
punteggio di 120 punti, un ottavo di un buco, quindi qualsiasi altro vincolo lo travolgeva. Ora il peso della
distanza sale da 40 a 200 e, soprattutto, è **simmetrico**: chi abita vicino paga le ore presto
(`PESO_VICINO_PRESTO = 150 × (1 − lontananza) × ore che mancano alla fine`), non solo il lontano paga le ore
tarde. Lo stesso scambio vale adesso circa 960 punti, appena sotto un buco: l'ordine per distanza si impone
su tutto tranne che sui buchi. In più l'attesa fra mattino e pomeriggio pesa in proporzione alla distanza
(× 0,3 per il più vicino invece che × 1): chi sta a dieci minuti può tornare a casa nel frattempo, e senza
questo correttivo non lo si sarebbe mai potuto spostare sul tardi. Sui dati veri, sui tre terzi per distanza:

| gruppo | ora media | lezioni alle 13:30 | lezioni alle 16:30 |
|--------|-----------|--------------------|--------------------|
| 25 più vicini | 1,85 | 9 | 23 |
| 25 di mezzo | 1,35 | 19 | 13 |
| 25 più lontani | 0,61 | 40 | 1 |

Degl'Innocenti passa a 2 rientri, Chen alle 14:30 e 15:30, buchi totali 1.

**Adattare invece di ricalcolare (15/9/2026)**: la casella «Parti dall'orario già calcolato» c'era già ed è
accesa da sola quando il file contiene un orario; il peso di spostare una lezione è però sceso da 1500 a
**800**, sotto il costo di un buco. Prima adattare l'orario alle regole nuove poteva lasciare dei vuoti pur
di non muovere le lezioni. Sui dati veri, adattando l'orario vecchio alle regole nuove: 264 lezioni
confermate su 283, 19 spostate, nessuna nuova.

**Conta l'ora in cui arriva a casa, non la durata del viaggio (14/9/2026, dall'analisi di un output)**:
Fiesoli, che sta a 128 minuti, si era vista mettere una lezione all'ultima ora. Guardando i suoi dati il
motivo era chiaro: finendo alle 16:30, alle 17:30 o alle 18:30 lei arriva **sempre alle 19:38**, perché il
treno è quello; il costo però era «minuti di attesa + viaggio», quindi finire più tardi *riduceva* l'attesa e
l'ultima fascia risultava la scelta migliore. Ora il costo è il **ritardo dell'arrivo a casa** rispetto al
meglio che quel ragazzo può fare (per lei: 17:03 finendo alle 14:30), moltiplicato per la lontananza. Le
fasce che portano allo stesso treno costano uguale, e a quel punto vince la più presta.

**I lontani saltano l'ultima ora e vengono un pomeriggio solo (14/9/2026, richiesta di Francesco)**: due pesi
nuovi, entrambi progressivi con la lontananza e senza soglie. `PESO_ULTIMA_LONTANO = 400 × lontananza` sulla
fascia delle 16:30, e `PESO_RIENTRO_LONTANO = 350 × lontananza` per ogni pomeriggio oltre il primo (prima un
secondo rientro era gratis fino al massimo consentito). Restano sotto il peso di un buco, quindi non si crea
un vuoto pur di anticipare. Sui dati veri: dei 30 studenti più lontani **uno solo** ha ancora una lezione alle
16:30 (erano molti di più), e Fiesoli è passata da giovedì 14:30-16:30 a lunedì 13:30-15:30.

**Gruppi di musica da camera negli abbinamenti fissi (14/9/2026, richiesta di Francesco)**: nella colonna
Studente si scrive «Gruppo 5» (vanno bene anche «Gr. 5», «LMC 5» o il solo numero) e l'ora del gruppo resta
lì: docente e tutti i membri occupati. Il docente si può lasciare vuoto, si prende dal foglio Gruppi LMC;
se se ne scrive uno diverso da quello del gruppo, il calcolo si ferma e lo dice. Sotto il cofano l'abbinamento
punta al primo membro del gruppo, così il motore riusa la strada già esistente: la lezione di gruppo è una
sola unità. Restano valide anche le due forme di prima: un singolo ragazzo (con tipo «Musica da camera»
si fissa comunque la lezione di tutto il suo gruppo) e i laboratori LMI.

**Riallineamento del formato (14/9/2026)**: `aggiorna_struttura` aggiunge fogli, colonne e parametri, ma non
tocca il *formato* del file: un file compilato mesi fa si ritrova le istruzioni vecchie e le colonne nuove
senza menu a tendina né larghezze. `template.rinfresca_formato(percorso)` riscrive il foglio Istruzioni e
rimette, foglio per foglio, larghezze, stile dell'intestazione e menu a tendina presi da un modello costruito
al momento con lo stesso numero di righe (così le tendine coprono tutte le righe vere). Un foglio le cui
colonne non corrispondono al modello viene lasciato stare. Si richiama dal menu **Modifica → Riallinea il
formato del file**.

**Colonna «Minuti per tornare a casa» (14/9/2026, richiesta di Francesco)**: nel foglio Studenti, dopo i KM.
La scrive il programma (il migliore delle 4 fasce, dal foglio Trasporti) quando si aggiornano i trasporti e a
ogni salvataggio, così il tempo di viaggio si legge accanto al ragazzo senza aprire il foglio Trasporti, dove
resta il dettaglio ora per ora. Non si compila a mano: viene riscritta.

**Colonna «1° strumento attaccato» (14/9/2026, richiesta di Francesco)**: sostituisce «Ore consecutive».
**SI** = le 2 ore di 1° strumento una di seguito all'altra (come prima); **NO o casella vuota** = in due
giorni diversi, mai attaccate. Il comportamento normale è quindi la separazione: prima le due ore finivano
quasi sempre nello stesso pomeriggio perché così non si lasciano buchi. Vale solo per le classi 1ª, 2ª e 5ª,
le uniche con 2 ore; nelle altre è un avviso e viene ignorata. Due eccezioni, entrambe segnalate: **le ore
fissate nel foglio degli abbinamenti** restano dove le ha messe la persona, e **«Giorno unico = SI»** vince
(non si può venire un giorno solo e avere le due ore in giorni diversi). I controlli si accorgono prima se la
richiesta è impossibile: per il SI servono due ore consecutive libere dello stesso docente, per la
separazione almeno due giorni. Aprendo un file di prima, la colonna viene rinominata mantenendo i SI scritti.
Sui dati veri: 54 ragazzi su 61 con due ore separate, i 7 restanti sono quelli fissati a mano.

**Laboratori LMI negli abbinamenti fissi (14/9/2026, richiesta di Francesco)**: nella colonna Studente si
può scrivere, al posto di un ragazzo, il nome di un laboratorio del foglio LMI (tipo «Laboratorio LMI»). Il
programma occupa in quell'ora il docente e tutti i ragazzi del laboratorio, e la lezione compare nell'orario
con il nome del laboratorio e l'elenco dei partecipanti. Il docente si può lasciare vuoto: si prende quello
scritto nel foglio LMI. I laboratori restano fuori dal calcolo (sono al mattino) finché non vengono fissati
così: è l'unico modo per portarli nel pomeriggio. Se un ragazzo del laboratorio ha un impegno in quell'ora, o
se il laboratorio non esiste, il calcolo si ferma e lo dice; i nomi del foglio LMI che non corrispondono a
nessuno studente diventano un avviso e quel ragazzo semplicemente non viene tenuto occupato.

**Le due ore di 1° strumento non sono attaccate (14/9/2026, richiesta di Francesco)**: non lo sono mai state
per regola (solo «Ore consecutive = SI» le lega), ma il foglio degli abbinamenti non permetteva di separarle
davvero: fra due ore uguali il modello imponeva un ordine per rompere la simmetria, e fissando a mano la
seconda prima della prima il calcolo diventava impossibile. Ora quella regola di simmetria salta appena una
delle due ore è fissata: si possono mettere dove si vuole, anche in giorni diversi e in ordine sparso. Se le
si vogliono di seguito, si scrivono due righe su due ore consecutive.

**Griglia degli abbinamenti più larga (14/9/2026)**: le colonne si adattavano al testo, e con la freccina del
menu a tendina restava mezza cella leggibile. Ora i fogli con i menu (Abbinamenti fissi, Gruppi LMC, LMI)
hanno una larghezza minima per colonna.

**Foglio «Abbinamenti fissi» (14/9/2026, richiesta di Francesco)**: la tabella dove si decidono a mano le
lezioni già fissate — docente, studente, tipo di lezione, giorno, ora, note. Il motore le blocca prima di
tutto il resto (`_blocca_abbinamenti`) e costruisce l'orario intorno. Le colonne Docente, Studente, Tipo di
lezione, Giorno e Ora hanno il menu a tendina con i valori possibili, e lo studente si può scrivere col solo
cognome come nei gruppi. Se un abbinamento è impossibile (docente non disponibile a quell'ora, studente
impegnato, due lezioni sulla stessa cella) il calcolo si ferma e lo dice.

**Somma delle ore in fondo alle disponibilità (14/9/2026, richiesta di Francesco)**: il foglio Docenti ha una
colonna «Ore dichiarate» (le X della riga) e in fondo una riga **TOTALE** con la somma per ogni fascia e il
totale generale. Si aggiornano da sole a ogni salvataggio e col pulsante «Aggiorna totali»; la riga TOTALE non
viene letta come un docente e resta sempre una sola.

**Aggiornamento dei file delle versioni precedenti (14/9/2026, richiesta di Francesco)**: aprendo un file
creato con una versione più vecchia, `aggiorna_struttura` lo porta alla struttura corrente — aggiunge i fogli
mancanti (Impegni studenti, Abbinamenti fissi…) nella posizione giusta, le colonne mancanti in coda a ogni
foglio (per esempio le 5 aule al posto dell'aula unica, o «Ore dichiarate») e le righe di parametro nuove,
senza toccare i dati già scritti. Prima di modificare fa una copia «… - prima dell'aggiornamento.xlsx» e alla
fine elenca che cosa ha cambiato. Il confronto dei titoli di colonna è esatto a meno delle parentesi tipo
«(SI/NO)»: con il confronto per prefisso «Ore dichiarate» veniva scambiata per «Ore accompagnamento».

**Riepilogo delle approssimazioni prima del calcolo (14/9/2026, richiesta di Francesco)**: premendo
Calcola il programma esegue subito i controlli e, se ha dovuto dare qualcosa per buono, apre una finestra
che lo elenca raggruppato per tipo, con "Calcola comunque" e "Annulla, correggo i dati". Le categorie sono
quelle decise in queste settimane: ore aperte d'ufficio a un docente, ragazzi senza distanza né tempi dei
mezzi, indirizzi non trovati con i tempi presi dal centro del comune, tempi dei mezzi non ancora calcolati,
nomi riconosciuti per somiglianza, righe con la nota ESEMPIO rimasta, impegni riferiti a ragazzi non in
elenco, richieste di ore consecutive ignorate, docenti senza margine e docenti senza lezioni. Ogni avviso
porta con sé la sua `categoria` (campo nuovo di `Problema`).

**Indirizzi: si usa la via, non il comune**: la geolocalizzazione parte sempre da via e civico e scende al
comune solo se le mappe non conoscono quella via. Sul file vero: 100 indirizzi su 103 risolti con la via
esatta, 2 dal centro del comune, 1 senza percorso. I due approssimati compaiono nel riepilogo.

**Esportazione in Word (14/9/2026, richiesta di Francesco)**: oltre a Excel e PDF il programma scrive i tre
documenti anche in `.docx` (`orario/export_docx.py`, con python-docx), riusando le stesse celle di PDF ed
Excel: settimanale orizzontale con una pagina per giorno, per docente e per studente in verticale, grassetto
per il 1° strumento, corsivo per il 2°, celle grigie dove il docente non è disponibile, intestazioni ripetute
se una tabella si spezza. Servono per ritoccare l'orario a mano prima di stamparlo.

**Legenda delle caselle vuote (14/9/2026)**: in fondo a ogni pagina ora c'è scritto che una casella bianca
vuota vuol dire "docente disponibile ma senza lezione in quell'ora", mentre la casella grigia vuol dire
"docente non disponibile". Era una domanda ricorrente: per esempio Falchi dichiara 10 ore ma ha un solo
allievo di corno, quindi la sua colonna è quasi tutta bianca e vuota.

**Foglio «Impegni studenti» (14/9/2026, richiesta di Francesco)**: una griglia come quella dei docenti ma
**al contrario**: la X segna l'ora in cui il ragazzo NON può esserci (sport, catechismo, altro). Un rigo per
ragazzo con classe, cognome e nome; chi non ha impegni resta vuoto. È un vincolo rigido, si somma alla
colonna "Giorni NON disponibili" del foglio Studenti (che vieta il giorno intero). Nel programma il pulsante
"Copia i nomi dagli studenti" riempie l'elenco. I file creati prima non hanno il foglio: viene aggiunto vuoto
all'apertura e scritto nel file al primo salvataggio.

**Attesa dopo il mattino (14/9/2026, dall'analisi dell'output)**: il programma contava i buchi *fra* le
lezioni del pomeriggio ma non l'attesa fra la fine del mattino e la prima lezione, che per chi resta a scuola
è tempo perso uguale. Era il motivo per cui capitava una lezione isolata alle 15:30 o alle 16:30. Ora ogni
ora di attesa costa 120, moltiplicato per la lontananza del ragazzo. Sui dati veri: ore di attesa da 128 a
80, pomeriggi che cominciano alle 16:30 da 19 a 5, ragazzi che vengono un solo pomeriggio da 50 a 65, sempre
con zero buchi.

**Fasce aperte d'ufficio (14/9/2026, richiesta di Francesco)**: se a un docente le ore dichiarate non
bastano per gli studenti che gli sono stati assegnati, il calcolo non si ferma più. Il programma apre le
fasce mancanti fra quelle segnate come non disponibili, scegliendole dove disturbano meno (prima nei giorni
in cui il docente c'è già, poi accanto alle fasce che ha, poi nelle ore più presto) e lo scrive negli
avvisi con l'elenco preciso, «da concordare con il docente». Resta un errore solo se le ore richieste
superano le 20 fasce della settimana, che nessuna apertura può risolvere.

**Nomi con qualche lettera diversa (14/9/2026)**: nei gruppi un nome scritto quasi come nel foglio Studenti
(«Valentini Niccolo Gairo» contro «Valentini Nicolo' Giairo») non blocca più il calcolo: viene riconosciuto
per somiglianza e segnalato, così si può correggere con calma.

**Indirizzi: vie sbagliate scartate (14/9/2026)**: Photon risponde "a somiglianza" e su un indirizzo che
non esiste nelle mappe restituiva comunque una via qualsiasi dello stesso comune (per «Via Renaggio» a
Montecatini dava «Via Ugolino da Montecatini», per «De Papiglioni» all'Abetone «Via Del Libro Aperto»). Ora
il nome della via trovata viene confrontato con quello chiesto, ignorando "Via/Piazza/Località" e le
paroline, e con le iniziali puntate che non fanno testo; se non somiglia, il risultato si scarta e si scende
al ripiego successivo, fino al centro del comune. Meglio un punto approssimativo ma vero che una via
sbagliata: l'esito lo dichiara, «OK (via non trovata: usato il centro del comune)».

**Righe di esempio (9/9/2026)**: il controllo bloccava il calcolo appena trovava la parola «ESEMPIO» nella
colonna Note, anche quando la riga era stata riempita con dati veri e restava solo la nota. Ora la riga è
considerata di esempio soltanto se contiene ancora i nomi inventati del modello (Rossi Mario, Neri Anna,
Bianchi, Verdi); altrimenti si ottiene un avviso che suggerisce di cancellare la nota. Nello stesso giro:
la lettura del file non si interrompe più quando trova solo avvisi, che ora viaggiano dentro `DatiInput` e
compaiono insieme a quelli dei controlli.

**Difetto grave corretto (9/9/2026)**: applicando il menu a tendina dei nomi, la libreria delle griglie
scriveva il primo valore dell'elenco in **tutte** le celle degli studenti dei gruppi (`edit_data=True` è il
suo comportamento predefinito). Ora il menu si applica con `edit_data=False` e come opzione di colonna, non
cella per cella: non tocca il contenuto ed è molto più leggero. Inoltre il completamento del nome agisce solo
sulla cella appena scritta invece di riscrivere tutto il foglio, la riga libera in fondo non viene inserita
mentre un editor di cella è aperto e la selezione viene ripristinata dopo l'inserimento: erano queste tre
cose a far perdere il fuoco della cella durante la scrittura.

**Griglie come in Excel (9/9/2026)**: copia, taglia, incolla, annulla e ripeti con le scorciatoie di
sistema; incollando più righe di quelle presenti le righe si aggiungono (fino a 5000); Invio scende, Tab va
a destra; una riga vuota è sempre pronta in fondo e scrivendoci dentro ne compare un'altra (le righe vuote
non vengono salvate); ricerca e sostituzione; tasto destro per inserire ed eliminare righe; righe
trascinabili e ordinabili per colonna; ridimensionamento di righe e colonne col mouse o doppio clic;
ingrandimento con ⌘/Ctrl e rotella o gesto del trackpad. Le colonne non si possono spostare, per non
disallineare i nomi delle intestazioni su cui si basa la lettura.

**Ordine delle colonne del foglio Studenti (9/9/2026, richiesta di Francesco)**: Classe, Cognome, Nome,
Strumento 1, Docente 1, Strumento 2, Docente 2, poi le eccezioni (ore consecutive, giorno unico, giorni non
disponibili), poi i dati di casa (Comune, Indirizzo, Civico, KM) e infine le Note. La lettura del file è per
nome di colonna, quindi i file già compilati con l'ordine vecchio continuano a funzionare.

**Test automatici (9/9/2026)** — 96 test, tutti verdi:
- `test_e2e.py` (74): template e lettura/scrittura del file, tutti i messaggi dei controlli, invarianti del
  motore riverificate in modo indipendente (ore per classe, sovrapposizioni, A e accompagnamento in coda,
  gruppi in coincidenza, giorno unico, ore consecutive, rientri), casi impossibili, ricalcolo, export
  Excel e PDF (con `pdftotext`/`pdfinfo`), trasporti con le risposte del servizio simulate.
- `test_gui.py` (22): apertura, nuovo file, griglie, menu, calcolo con scelta della cartella, finestra dei
  problemi, pulsante Trasporti e flusso concatenato trasporti → orario, robustezza.
- `test_motore.py`, `test_export.py`: prove storiche sul dataset grande (109 studenti).

**Bug trovati dai test e corretti (9/9/2026)**
- "Chiudi file" e "Cambia file" non chiudevano davvero il file: lo stato restava in memoria e le voci
  Salva/Calcola/Trasporti continuavano a lavorare su griglie non più visibili (Calcola arrivava a
  esportare). Ora la schermata dati viene distrutta e lo stato azzerato.
- `_aggiorna_info_trasporti` poteva configurare un pulsante già distrutto (è schedulata con `after`).
- `nuovo_file` poteva far crollare il programma se il file appena creato non era leggibile.
- La cartella dei risultati aveva la data al minuto: due calcoli nello stesso minuto si sovrascrivevano.
  Ora il nome include i secondi.
- `motore.calcola` chiamato senza `controlli.controlla` produceva gruppi di musica da camera vuoti e un
  errore interno fuorviante: ora si ferma con un messaggio esplicito.
- Tolto un controllo che non poteva mai scattare (giorno unico con più di 4 ore settimanali).

**Note tecniche emerse dalle prove (8/9/2026)**
- Sui dati di prova (109 studenti, 23 docenti, 284 lezioni) il motore trova in 5 s una soluzione buona e in
  40–120 s una con **zero buchi** per studenti e docenti; quasi tutti con 2 rientri, nessuno con 3. La
  distanza progressiva funziona: chi abita oltre 25 km ha in media lezione alla 7ª ora, chi abita sotto 5 km
  alla 9ª.
- Ricalcolo dopo una modifica (tolta una disponibilità a un docente): meno di un secondo, 2 lezioni spostate
  su 284, con spiegazione negli avvisi. Peso dello spostamento: preferisce spostare due lezioni piuttosto che
  lasciare a un ragazzo due ore di buco.
- **Attenzione alla A**: una A fissata nella prima fascia del giorno impedisce al docente qualunque lezione
  quel giorno (l'accompagnamento deve stare in coda). Va usata solo nelle ultime fasce.
- Il PDF settimanale sta su una pagina per giorno fino a ~25 docenti; oltre, il carattere diventa troppo
  piccolo e servirà spezzare su due pagine.
- L'anno nel titolo del PDF viene preso dal nome del file di input se contiene "2026-27" (o simile).

---

## 10. Come si usa

**Da terminale (oggi)**
```bash
# una volta sola
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# avvio del programma
.venv/bin/python app.py

# (facoltativo) dati inventati per provare
.venv/bin/python dati_prova.py
```

**Documentazione**: `README.md` (panoramica), `WINDOWS.md` (come creare il programma per Windows, per
chi sviluppa), `GUIDA INSTALLAZIONE WINDOWS.md` (guida per l'utente finale: installazione, uso, problemi
tipici), questo file (progettazione e decisioni).

**Con doppio clic su Mac**: `./costruisci_app.sh` crea `dist/Orario Musicale.app`.

**Su Windows** (vedere `WINDOWS.md` per i dettagli): il `.exe` va costruito **su** Windows, dal Mac non si può.
Tre strade: `costruisci_installer.bat` (consigliata per la consegna: PyInstaller + **Inno Setup**, produce
`installer_output\OrarioMusicale-setup-1.0.exe`, installazione per singolo utente senza diritti di
amministratore, collegamenti e disinstallazione; lo script installa Inno Setup con winget se manca);
`windows_avvia.bat` (installa l'ambiente al primo avvio e lancia il programma da sorgente, serve Python con
"Add python.exe to PATH"); `costruisci_app.bat` (solo la cartella `dist\Orario Musicale`, da copiare
intera). Al primo avvio
SmartScreen chiede "Ulteriori informazioni" → "Esegui comunque"; gli antivirus a volte mettono in quarantena
gli eseguibili creati con PyInstaller, in quel caso conviene la prima strada.

**Accorgimenti per Windows già presi**: `tzdata` nei requisiti (Windows non ha il database dei fusi, che
serve per gli orari dei mezzi) con calcolo di riserva dell'ora legale italiana se mancasse; carattere Segoe UI
nelle griglie; salvataggio del file Excel con alcuni tentativi, perché antivirus o Excel possono tenerlo
occupato per un istante.

**Flusso**: Apri file Excel (o Nuovo file) → correggi nelle griglie → Salva → Calcola orario → scegli dove
salvare → cartella "Risultati orario <data e ora>" con `orario.xlsx`, `orario_settimanale.pdf`,
`orario_docenti.pdf`, `orario_studenti.pdf`. Se i dati non tornano, compare l'elenco dei problemi e nessun file viene scritto.
Al calcolo successivo, la casella "Parti dall'orario già calcolato" mantiene l'orario e sposta solo il necessario.
