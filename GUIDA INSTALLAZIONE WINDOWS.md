# Orario Musicale — installazione su Windows

Guida per chi riceve il programma. Serve una decina di minuti, una volta sola.

---

## 1. Installare il programma

1. Ti arriva un file che si chiama **`OrarioMusicale-setup-1.0.exe`** (via chiavetta, email o link).
   Salvalo dove vuoi, per esempio sul Desktop o in Download.
2. **Doppio clic** sul file.
3. Quasi certamente comparirà una finestra blu: **"Windows ha protetto il PC"**. È normale: succede a
   tutti i programmi che non hanno una firma digitale a pagamento. Fai così:
   - clicca su **"Ulteriori informazioni"** (la scritta piccola);
   - poi sul pulsante **"Esegui comunque"** che appare sotto.

   Questa schermata compare solo la prima volta.
4. Parte l'installazione. Vai avanti con **Avanti** fino alla fine. Lascia spuntato
   "Crea un'icona sul desktop" se la vuoi.
5. Alla fine spunta **"Avvia Orario Musicale"** e clicca **Fine**.

Il programma non chiede la password di amministratore e si installa nella tua cartella personale, quindi
non serve chiamare nessuno.

Da ora in poi lo apri con l'icona sul desktop, oppure cercando "Orario" nel menu Start.

---

## 2. Il file di esempio

Durante l'installazione viene messo un file di prova qui:

```
Documenti\Orario Musicale\orario_2026-27.xlsx
```

Contiene 30 ragazzi inventati, 7 docenti e sei gruppi di musica da camera: serve per prendere la mano
senza toccare i dati veri. Aprilo dal programma e prova a calcolare: ci vuole qualche secondo.

---

## 3. Come si usa, in breve

1. Apri il programma. Clicca **Apri file Excel…** e scegli il tuo file. Se devi partire da zero clicca
   **Nuovo file…**: te ne crea uno pronto da compilare (le righe grigie con scritto "ESEMPIO" sono lì
   per far vedere come si riempie, poi vanno cancellate).
2. Vedi cinque schede in alto: **Studenti, Docenti, Gruppi LMC, LMI, Parametri**. Sono come i fogli di
   Excel: doppio clic su una cella per modificarla, tasto destro per aggiungere o togliere righe.
3. Quando hai finito clicca **💾 Salva**.
4. Clicca **⚡ Calcola orario**. Il programma chiede in quale cartella salvare i risultati e lì crea una
   cartella con la data di oggi, contenente:
   - `orario.xlsx` — l'orario in Excel, un foglio per ogni giorno
   - `orario_settimanale.pdf` — la griglia da stampare e appendere
   - `orario_docenti.pdf` — una pagina per ogni professore
   - `orario_studenti.pdf` — le lezioni di ogni ragazzo, da dare alle famiglie
5. Se nei dati c'è qualcosa che non torna, il programma **si ferma** e ti mostra l'elenco preciso dei
   problemi, riga per riga, senza scrivere nessun file. Correggi e riprova.

---

## 4. Il pulsante 🚌 Trasporti

Calcola, per ogni ragazzo, quanto tempo ci vuole per tornare a casa con bus e treni se la lezione finisce
alle 14:30, alle 15:30, alle 16:30 o alle 17:30. Poi il programma mette in orario **prima** chi ha il
viaggio più lungo ed evita le ore dopo le quali non c'è più un mezzo.

- Serve **internet**, e ci vogliono circa 3 secondi per ragazzo (per un centinaio di ragazzi, cinque
  minuti). Si fa una volta e basta: i risultati restano nel file.
- Perché funzioni, nel foglio **Studenti** devono essere compilate le colonne **Comune, Indirizzo,
  Civico**, e nel foglio **Parametri** la riga **Indirizzo della scuola**.
- Il pulsante ti dice come stai: scritto in blu "60/80" significa che a 20 ragazzi mancano i tempi; verde
  con "✓ 80" significa che sono tutti a posto.
- Le volte successive ti chiede se rifare tutti (serve solo quando cambiano gli orari dei mezzi, di solito
  a settembre) o solo i ragazzi nuovi: in quel caso ci mette pochi secondi.
- Se un indirizzo non viene trovato, il programma te lo dice e per quel ragazzo usa i chilometri.

---

## 5. Ricalcolare dopo una modifica

Se cambi qualcosa (la disponibilità di un professore, un ragazzo in più) e ricalcoli, la casella in alto
**"Parti dall'orario già calcolato"** fa in modo che il programma **sposti solo le lezioni strettamente
necessarie**, lasciando tutto il resto dov'era. Alla fine, nella scheda **Avvisi**, trovi l'elenco di cosa
è stato spostato e perché.

Se invece vuoi ricominciare da zero, togli la spunta a quella casella (o usa il menu **Orario → Ricalcola
da zero**).

---

## 6. Se qualcosa va storto

**"Impossibile sovrascrivere il file: è aperto in un altro programma."**
Il file Excel è aperto in Excel. Chiudilo e clicca di nuovo Salva.

**L'antivirus blocca il programma o lo mette in quarantena.**
È un falso allarme, frequente con i programmi fatti così. Bisogna aggiungere un'eccezione per la cartella
del programma, oppure farsi dare la versione "cartella" invece dell'installer.

**Il programma non si apre e non dice niente.**
Riavvia il computer e riprova. Se continua, riferisci il problema a chi ti ha dato il programma.

**"Aggiorna trasporti" dà errore di rete.**
Manca internet, oppure la rete della scuola blocca le connessioni. Il programma continua a funzionare
usando i chilometri: gli orari dei mezzi si possono calcolare più tardi, anche da casa.

**Ho sbagliato tutto e voglio ripartire.**
I tuoi dati stanno solo nel file Excel. Fanne una copia prima di modifiche importanti: se qualcosa va
storto, basta tornare alla copia.

---

## 7. Disinstallare

Impostazioni di Windows → App → App installate → cerca "Orario Musicale" → Disinstalla.
I tuoi file Excel non vengono toccati.
