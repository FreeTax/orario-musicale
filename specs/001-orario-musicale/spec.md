# Feature Specification: Orario Pomeridiano Liceo Musicale

**Feature Branch**: `001-orario-musicale`

**Created**: 2026-06-29

**Status**: Draft

**Input**: User description: App desktop per la gestione automatica dell'orario
pomeridiano di un liceo musicale, con importazione dati da CSV, generazione
con vincoli hard/soft, esportazione PDF e CSV, e gestione multi-anno.

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Genera e scarica l'orario settimanale (Priority: P1)

Il coordinatore ha già caricato i dati dell'anno scolastico (studenti, gruppi
LMC, disponibilità docenti). Apre la schermata "Genera orario", avvia la
generazione, visualizza in anteprima la griglia risultante e la esporta in PDF
e CSV pronti per la distribuzione.

**Why this priority**: Questo è il motivo d'essere dell'app. Senza la
generazione automatica, tutto il resto non ha valore. È anche il flusso più
complesso e rischia di fallire se i vincoli sono incompatibili.

**Independent Test**: Con i tre CSV di esempio precaricati, cliccare "Genera
orario" → verificare che la griglia rispetti tutti i vincoli hard → esportare
PDF → aprire il file e verificare che sia leggibile e stampabile.

**Acceptance Scenarios**:

1. **Given** tutti i dati dell'anno corrente sono caricati, **When** il
   coordinatore clicca "Genera orario", **Then** il sistema produce una griglia
   valida in cui ogni studente ha il numero corretto di ore settimanali, ogni
   docente non è mai doppiamente prenotato e tutti i vincoli hard sono
   soddisfatti.

2. **Given** una griglia valida è stata generata, **When** il coordinatore
   clicca "Esporta PDF", **Then** viene salvato un file PDF con colonne =
   docenti e righe = fasce orarie, con cognome e classe dello studente in ogni
   cella occupata.

3. **Given** una griglia valida è stata generata, **When** il coordinatore
   clicca "Esporta CSV", **Then** viene salvato un file CSV con la stessa
   struttura del PDF (colonne = docenti, righe = fasce orarie).

4. **Given** i dati forniti rendono impossibile soddisfare tutti i vincoli
   hard, **When** il coordinatore avvia la generazione, **Then** il sistema
   mostra un messaggio di errore che indica quale vincolo non è soddisfacibile,
   senza produrre output parziali.

5. **Given** una griglia generata con vincoli soft parzialmente ottimizzati,
   **When** il coordinatore visualizza l'anteprima, **Then** la griglia mostra
   esplicitamente le fasce non assegnate e i "buchi" (fasce di attesa) per
   studente.

---

### User Story 2 - Importa e gestisci i dati di un anno scolastico (Priority: P2)

Il coordinatore crea o apre un anno scolastico e carica i tre file CSV
(studenti, gruppi LMC, disponibilità docenti). Può modificare i dati
direttamente nelle tabelle dell'app e riesportarli come CSV.

**Why this priority**: Senza dati corretti non è possibile generare l'orario.
La gestione dei dati è il prerequisito diretto della user story P1, ma è
indipendentemente testabile e consegnabile come schermata funzionante.

**Independent Test**: Creare un nuovo anno scolastico "2026-27" → importare
i tre CSV → verificare che tutte le righe siano visibili e corrette nelle
rispettive schermate → modificare un record direttamente in tabella → esportare
il CSV → verificare che il file esportato rifletta la modifica.

**Acceptance Scenarios**:

1. **Given** un nuovo anno scolastico viene creato dalla schermata Home,
   **When** il coordinatore importa `studenti.csv`, **Then** tutti gli studenti
   appaiono nella tabella con campi cognome, nome, classe, km, docente_1str,
   strumento_1, docente_2str, strumento_2, ore_consecutive, giorno_unico
   correttamente popolati.

2. **Given** la tabella studenti è visualizzata, **When** il coordinatore
   modifica direttamente un campo (es. cambia classe da 2 a 3), **Then** la
   modifica viene salvata e persiste alla chiusura e riapertura dell'app.

3. **Given** i dati sono stati modificati, **When** il coordinatore clicca
   "Esporta CSV" su qualsiasi schermata dati, **Then** il file CSV esportato
   contiene le modifiche più recenti e rispetta lo schema delle colonne
   originale.

4. **Given** la griglia delle disponibilità docenti è visualizzata,
   **When** il coordinatore clicca su una cella, **Then** alterna lo stato
   disponibile/non-disponibile per quella fascia oraria di quel docente.

5. **Given** un file CSV con dati non validi (colonne mancanti, valori fuori
   range) viene importato, **When** l'importazione viene avviata, **Then** il
   sistema mostra un messaggio di errore che indica le righe problematiche e
   non sovrascrive i dati esistenti.

---

### User Story 3 - Duplica anno scolastico con promozione automatica (Priority: P3)

Il coordinatore crea il nuovo anno scolastico duplicando quello precedente.
Il wizard promuove automaticamente gli studenti alla classe successiva, rimuove
i diplomati di quinta, svuota la prima classe, azzera i gruppi LMC e mantiene
le disponibilità docenti invariate.

**Why this priority**: Riduce significativamente il lavoro annuale di
re-inserimento dati. Non è bloccante per P1 e P2 (si può lavorare con dati
caricati manualmente), ma è critico per l'adozione a lungo termine.

**Independent Test**: Con un anno "2025-26" esistente contenente studenti di
tutte le classi, avviare "Duplica anno" → specificare "2026-27" →
verificare nel nuovo anno: classe 5 rimossa, classi 1-4 promosse, classe 1
vuota, gruppi LMC vuoti, disponibilità docenti invariate.

**Acceptance Scenarios**:

1. **Given** un anno scolastico esistente viene selezionato per la
   duplicazione, **When** il wizard viene completato con il nuovo label (es.
   "2026-27"), **Then** viene creata una nuova cartella "2026-27" con studenti
   promossi: ogni studente di classe N diventa classe N+1, i diplomati di
   quinta sono rimossi, la classe 1 è vuota.

2. **Given** la duplicazione è completata, **When** il coordinatore apre
   l'anno duplicato, **Then** le disponibilità docenti sono identiche
   all'anno sorgente e i gruppi LMC sono completamente vuoti.

3. **Given** il wizard di duplicazione è aperto, **When** il coordinatore
   revisa il riepilogo delle modifiche (studenti promossi, rimossi, classe 1
   azzerata), **Then** può confermare o annullare prima che qualsiasi dato
   venga scritto.

4. **Given** un anno con lo stesso label esiste già, **When** il coordinatore
   tenta di duplicare con quel label, **Then** il sistema mostra un avviso e
   richiede conferma prima di sovrascrivere.

---

### Edge Cases

- Cosa succede se un docente non ha disponibilità sufficienti per coprire
  tutte le lezioni degli studenti assegnati?
- Cosa succede se un gruppo LMC include studenti con `giorno_unico=SI` ma
  con giorni preferiti incompatibili?
- Cosa succede se la griglia oraria viene rieseguita dopo aver modificato i
  dati — il risultato precedente viene sovrascritto automaticamente?
- Cosa succede se `studenti.csv` contiene studenti di classe 5 con campi
  docente_2str o strumento_2 non vuoti (dati incongruenti)?
- Cosa succede se un gruppo LMC contiene uno studente che non è presente
  nell'elenco studenti (riferimento orfano)?

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: Il sistema DEVE consentire al coordinatore di creare un nuovo
  anno scolastico identificato da un'etichetta nel formato "AAAA-AA" (es.
  "2026-27") e archiviato in una cartella dedicata.

- **FR-002**: Il sistema DEVE consentire di aprire un anno scolastico
  esistente dalla schermata Home.

- **FR-003**: Il sistema DEVE consentire l'importazione di studenti da un
  file `studenti.csv` con le colonne: cognome, nome, classe (1-5), km,
  docente_1str, strumento_1, docente_2str (vuoto per classe 5),
  strumento_2 (vuoto per classe 5), ore_consecutive (SI/NO),
  giorno_unico (SI/NO).

- **FR-004**: Il sistema DEVE consentire l'importazione di gruppi LMC da un
  file `gruppi_lmc.csv` con le colonne: numero_gruppo, docente,
  studente_1, studente_2, studente_3 (opzionale), studente_4 (opzionale).
  I gruppi LMC DEVONO contenere da 2 a 4 studenti.

- **FR-005**: Il sistema DEVE consentire l'importazione delle disponibilità
  docenti da un file `disponibilita_docenti.csv` con 20 colonne di fasce
  orarie (lunedì–venerdì × 4 fasce: 13:30, 14:30, 15:30, 16:30). Il
  valore "X" indica disponibilità; la cella vuota indica non-disponibilità.

- **FR-006**: Il sistema DEVE consentire la modifica diretta di qualsiasi
  record studente, gruppo LMC o disponibilità docente nell'interfaccia
  senza dover reimportare il CSV.

- **FR-007**: Il sistema DEVE consentire l'esportazione di studenti, gruppi
  LMC e disponibilità docenti come file CSV con lo stesso schema di importazione.

- **FR-008**: Il sistema DEVE assegnare il numero corretto di lezioni
  settimanali per studente in base alla classe:
  - Classe 1-2: 2 ore 1° strumento + 1 ora 2° strumento
  - Classe 3-4: 1 ora 1° strumento + 1 ora 2° strumento + 1 ora LMC
  - Classe 5: 2 ore 1° strumento + 1 ora LMC

- **FR-009**: Il generatore di orario DEVE rispettare i seguenti vincoli
  hard (non derogabili):
  - Le lezioni vengono assegnate solo nelle fasce in cui il docente ha la "X"
  - Nessun docente può avere due lezioni nella stessa fascia oraria
  - Tutti i membri di un gruppo LMC DEVONO essere assegnati alla stessa
    fascia oraria dello stesso giorno
  - Le lezioni individuali dei membri di un gruppo LMC DEVONO essere
    collocate in fasce adiacenti dello stesso giorno in cui cade la LMC

- **FR-010**: Il generatore DEVE ottimizzare i vincoli soft nel seguente
  ordine di priorità:
  1. Minimizzare i buchi (fasce di attesa) tra lezioni dello stesso studente
     nello stesso pomeriggio
  2. Minimizzare i giorni di rientro per studente (max 2 giorni/settimana;
     max 3 giorni per studenti con km < 5)
  3. Collocare gli studenti con km più elevato nelle prime fasce della
     giornata (13:30, 14:30) e quelli con km più basso nelle ultime
     (15:30, 16:30)
  4. Rispettare ore_consecutive=SI: le 2 ore di 1° strumento devono essere
     in fasce adiacenti dello stesso giorno
  5. Rispettare giorno_unico=SI: tutte le lezioni dello studente nello
     stesso giorno

- **FR-011**: Il sistema DEVE posizionare i gruppi LMC prima di qualsiasi
  altra lezione nella sequenza di generazione.

- **FR-012**: Il sistema DEVE mostrare un'anteprima della griglia generata
  prima dell'esportazione, con colonne = docenti, righe = fasce orarie
  (giorno + orario) e celle = cognome studente + classe.

- **FR-013**: Il sistema DEVE esportare la griglia come PDF con colonne =
  docenti, righe = fasce orarie, celle = cognome studente + classe.

- **FR-014**: Il sistema DEVE esportare la griglia come CSV con la stessa
  struttura del PDF.

- **FR-015**: Il sistema DEVE mostrare un errore esplicito quando la
  generazione non riesce a soddisfare i vincoli hard, indicando la causa
  (es. docente senza disponibilità sufficienti), senza produrre output parziali.

- **FR-016**: Il sistema DEVE fornire un wizard "Duplica anno" che:
  - Promuove ogni studente dalla classe N alla classe N+1
  - Rimuove tutti gli studenti di classe 5
  - Svuota la lista studenti di classe 1
  - Azzera i gruppi LMC
  - Mantiene le disponibilità docenti invariate
  - Mostra un riepilogo delle modifiche prima di confermare

- **FR-017**: Il sistema DEVE impedire la sovrascrittura accidentale di un
  anno scolastico esistente durante la duplicazione, richiedendo conferma
  esplicita se l'etichetta esiste già.

### Key Entities

- **AnnoScolastico**: etichetta (es. "2026-27"), cartella di archiviazione
  contenente i tre file CSV di input e l'eventuale orario generato.

- **Studente**: cognome, nome, classe (1–5), km dalla scuola, docente e
  strumento per il 1° strumento, docente e strumento per il 2° strumento
  (vuoti per classe 5), flag ore_consecutive, flag giorno_unico.

- **GruppoLMC**: numero identificativo, docente LMC, da 2 a 4 studenti
  (solo di classi 3, 4 o 5).

- **DisponibilitaDocente**: docente, matrice di disponibilità su 20 fasce
  orarie (5 giorni × 4 fasce: 13:30–14:30, 14:30–15:30, 15:30–16:30,
  16:30–17:30).

- **FasciaOraria**: giorno della settimana (lun–ven) + orario di inizio
  (13:30, 14:30, 15:30, 16:30).

- **OrarioGenerato**: griglia settimanale che associa ogni coppia
  (docente, fascia oraria) a uno studente assegnato (o vuoto).

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: Il coordinatore genera un orario completo e valido a partire
  dai dati caricati in meno di 2 minuti (incluso il tempo di elaborazione).

- **SC-002**: Il 100% delle griglie generate soddisfa tutti i vincoli hard;
  nessun orario con violazioni hard viene consegnato come output.

- **SC-003**: Il coordinatore importa i tre file CSV e dispone di dati
  pronti per la generazione in meno di 5 minuti.

- **SC-004**: Il wizard "Duplica anno" completa la promozione automatica in
  meno di 30 secondi senza richiedere re-inserimento manuale dei dati degli
  studenti promossi.

- **SC-005**: Il file PDF esportato è leggibile e stampabile su carta A4
  in formato orizzontale senza troncamenti.

- **SC-006**: L'app si avvia e raggiunge la schermata Home in meno di
  10 secondi su hardware standard (CPU dual-core, 4 GB RAM).

## Assumptions

- L'applicazione è usata da un singolo utente (il coordinatore) alla volta;
  non è previsto l'uso concorrente o multi-utente.
- Le fasce orarie sono sempre le 4 stesse ogni giorno (13:30, 14:30, 15:30,
  16:30) e non sono configurabili dall'utente.
- Il cognome è sufficiente per identificare uno studente nei report di output;
  nel caso di omonimi il coordinatore distingue aggiungendo un identificatore
  manuale nel campo cognome (es. "Rossi-A").
- Il numero di gruppi LMC è variabile; "25 gruppi" è la dimensione tipica
  per la scuola, non un limite hard imposto dall'app.
- Se il solver non trova soluzione ottimale entro 60 secondi, restituisce la
  migliore soluzione parziale trovata con indicazione esplicita dei vincoli
  soft non soddisfatti; i vincoli hard non vengono mai violati.
- I file CSV di input usano punto e virgola come separatore e codifica UTF-8.
- L'app non fornisce meccanismi di backup automatico; la sicurezza dei dati
  è responsabilità del coordinatore tramite la copia delle cartelle annuali.
- Il campo `km` è un numero intero o decimale; valori assenti o non numerici
  vengono trattati come 0 (fascia più vicina, slot tardivi).
- Gli studenti referenziati nei gruppi LMC DEVONO essere presenti nell'elenco
  studenti; un riferimento orfano causa un errore di validazione prima della
  generazione.
