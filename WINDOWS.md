# Far girare il programma su Windows

Tre strade, in ordine di comodità per chi riceve il programma:

| Serve a | Cosa lanciare | Risultato |
|---|---|---|
| Usarlo sul PC dove lo prepari | `windows_avvia.bat` | Si apre il programma (serve Python) |
| **Consegnarlo a qualcun altro** | `costruisci_installer.bat` | Un `setup.exe` da installare |
| Portarlo su chiavetta senza installare | `costruisci_app.bat` | Una cartella con il programma dentro |

Un avvertimento che vale per tutte: **il programma per Windows va creato su Windows**. Dal Mac non si può,
quindi questi passaggi vanno fatti una volta su un PC Windows.

---

## Strada 1: doppio clic, senza installare nulla a mano

1. Copia sul PC Windows tutta la cartella del progetto (via chiavetta, OneDrive, cartella condivisa: come
   preferisci). Non serve la cartella `.venv` né `dist`, se ci sono puoi anche cancellarle.
2. Installa Python, una volta sola: <https://www.python.org/downloads/windows/> → "Download Python".
   **Durante l'installazione spunta "Add python.exe to PATH"**, in basso nella prima schermata.
3. Doppio clic su **`windows_avvia.bat`**.
   - La prima volta prepara l'ambiente e scarica le librerie: qualche minuto, serve internet.
   - Le volte dopo apre il programma in pochi secondi.

Si apre una finestra nera del prompt insieme al programma: è normale, va lasciata aperta finché usi il
programma.

---

## Strada 2: creare un installer da consegnare (consigliata)

È la via più comoda se il programma lo deve usare qualcun altro: si ottiene un unico file `setup.exe`,
come quelli a cui siamo abituati, con schermate in italiano, collegamento sul desktop, voce nel menu Start
e disinstallazione dal Pannello di controllo.

1. Sul PC Windows, con Python installato come sopra, doppio clic su **`costruisci_installer.bat`**.
   Lo script fa tutto: ambiente Python, programma, installer. La prima volta installa anche **Inno Setup**
   (il programma gratuito che crea l'installer): se non riesce da solo, lo scarichi da
   <https://jrsoftware.org/isdl.php>, lo installi e rilanci lo script.
2. Alla fine trovi **`installer_output\OrarioMusicale-setup-1.0.exe`**: circa 100 MB, un file solo.
   Si consegna su chiavetta o con un link di condivisione (per email spesso è troppo grande).
3. Chi lo riceve fa doppio clic e segue le schermate. **Non serve Python e non serve la password di
   amministratore**: il programma si installa nella cartella personale dell'utente.
4. Un file Excel di esempio viene messo in **Documenti\Orario Musicale**.

Per aggiornare il programma in futuro basta rifare l'installer e lanciarlo: riconosce la versione già
installata e la sostituisce, senza toccare i file Excel dei dati.

---

## Strada 3: la cartella del programma, senza installer

Se preferisci non usare un installer, va bene anche così.

1. Sul PC Windows, con Python installato come sopra, doppio clic su **`costruisci_app.bat`**.
2. Alla fine trovi il programma in **`dist\Orario Musicale\Orario Musicale.exe`**.
3. **Copia tutta la cartella `dist\Orario Musicale`**, non solo il file `.exe`: dentro ci sono le librerie
   che servono. Puoi metterla dove vuoi, anche su una chiavetta.
4. Per comodità: tasto destro sul file `.exe` → "Invia a" → "Desktop (crea collegamento)".

La costruzione richiede qualche minuto e la cartella risultante pesa circa 300 MB.

---

## Cose che possono succedere la prima volta

**"Windows ha protetto il PC" (SmartScreen).** Compare perché il programma non ha una firma digitale a
pagamento. Clicca su **"Ulteriori informazioni"** e poi su **"Esegui comunque"**. Succede solo la prima volta.

**L'antivirus mette il file in quarantena.** Capita con i programmi creati in questo modo: è un falso
allarme. Bisogna aggiungere un'eccezione per la cartella, oppure usare la Strada 1, che non crea nessun
`.exe` e quindi non dà questo problema.

**Il programma non si apre e non dice niente.** Apri il prompt dei comandi nella cartella e lancia
`.venv\Scripts\python.exe app.py`: così l'eventuale errore si vede scritto.

**"Impossibile sovrascrivere il file: è aperto in un altro programma".** Il file Excel di input è aperto
in Excel. Chiudilo e premi di nuovo Salva.

**Antivirus e "Aggiorna trasporti".** Il calcolo dei mezzi ha bisogno di internet. Se l'antivirus o la
rete della scuola bloccano le connessioni, il programma lo dice e usa i chilometri.

---

## Differenze rispetto al Mac

Nessuna nel funzionamento. Sono stati sistemati tre punti che su Windows si comportano diversamente:

- **Fusi orari.** Windows non ha il database dei fusi che serve per leggere gli orari dei mezzi. Ora viene
  installato insieme alle librerie (`tzdata`) e, se mancasse, il programma calcola da sé l'ora legale
  italiana.
- **Caratteri.** Le griglie usano Segoe UI su Windows e Helvetica su Mac, così i testi non risultano
  sgranati.
- **Salvataggio.** Su Windows può succedere che un antivirus tenga il file occupato per un istante: il
  salvataggio riprova alcune volte prima di segnalare un problema.

---

## Aggiornare il programma in seguito

Se ti mando una versione nuova: sostituisci la cartella `orario` e i file `app.py`, `requirements.txt`,
poi rilancia lo script che hai usato (`windows_avvia.bat`, `costruisci_installer.bat` o
`costruisci_app.bat`). I file Excel con i dati non vengono toccati: tienili comunque fuori dalla cartella
del programma, per sicurezza.
