# Research: Orario Pomeridiano Liceo Musicale

**Feature**: `001-orario-musicale`
**Date**: 2026-06-29

## Decision 1: Modello CP-SAT per lo scheduling

**Decision**: OR-Tools CP-SAT con variabili booleane `x[lezione_id, giorno,
fascia]` = 1 se quella lezione è assegnata a quel giorno+fascia. Strategia
"LMC first": i gruppi LMC sono schedulati nella prima passata vincolando i
membri allo stesso slot, poi le lezioni individuali vengono collocate in fasce
adiacenti dello stesso giorno.

**Rationale**: CP-SAT è il solver più adatto per scheduling con vincoli misti
hard/soft. Il problema è un constraint satisfaction + optimization: ~200
studenti × 3 lezioni/studente × 20 fasce = ~12.000 variabili booleane, ampia-
mente entro i limiti pratici di CP-SAT. Il timeout di 60 s è configurabile con
`solver.parameters.max_time_in_seconds = 60`. Se scade, CP-SAT restituisce
la migliore soluzione trovata (o INFEASIBLE se nessuna soluzione è stata trovata).

**Alternatives considered**:
- Algoritmo greedy: più veloce ma non garantisce backtracking su violazioni soft;
  non può certificare infeasibilità hard
- Simulated annealing: adatto per ottimizzazione ma non per vincoli hard stretti

**Modello interno**:

```
Lezioni da schedulare (per ogni studente s, tipo t, indice i):
  x[s,t,i, d, f] ∈ {0,1}   d ∈ [0,4], f ∈ [0,3]

Vincoli hard:
  H1. Σ_{d,f} x[s,t,i,d,f] = 1                     — ogni lezione assegnata 1 volta
  H2. x[s,t,i,d,f] = 0 se docente(s,t) ∉ disponibilita[d][f]
  H3. Σ_{(s,t,i) con stesso docente} x[s,t,i,d,f] ≤ 1   — no doppia prenotazione
  H4. LMC: x[s1,lmc,d,f] = x[s2,lmc,d,f] per ogni coppia del gruppo
  H5. Membro LMC: la lezione individuale di s nel giorno d_lmc cade in f±1
      (fascia adiacente alla LMC)

Obiettivo soft (minimizzare somma pesata):
  S1 (peso 100): buchi = fasce vuote tra prima e ultima lezione di s in giorno d
  S2 (peso 50):  giorni_extra = max(0, giorni_rientro(s) - limite(s))
  S3 (peso 20):  dist_penalty: studenti lontani (km≥10) in fasce tardive (f≥2)
  S4 (peso 80):  ore_consecutive=SI violate: le 2 ore di 1str non adiacenti
  S5 (peso 80):  giorno_unico=SI violato: lezioni in più giorni
```

---

## Decision 2: Tabella editabile in customtkinter

**Decision**: `tkinter.ttk.Treeview` per la visualizzazione tabellare, con un
widget `CTkEntry` sovrapposto dinamicamente sulla cella in doppio-click per
l'editing inline. Binding Enter/FocusOut per confermare la modifica.

**Rationale**: customtkinter non ha un widget tabella nativo. `ttk.Treeview`
è parte della stdlib Python (zero dipendenze extra), supporta scroll, selezione
multipla, ordinamento colonne e risulta stabile su macOS e Windows. L'editing
inline tramite Entry sovrapposto è il pattern standard per le app Tkinter
desktop. Lo styling CTk non si propaga ai widget ttk nativi, ma l'aspetto è
accettabile in una app enterprise interna.

**Alternatives considered**:
- `CTkTable` (pacchetto third-party): dipendenza extra non giustificata, meno
  matura di ttk.Treeview
- `CTkScrollableFrame` + griglia di Entry: costruzione manuale pesante, resize
  problematico con molte righe
- Pandas DataFrame viewer custom: riduplica lavoro già fatto da ttk.Treeview

**Pattern per la griglia disponibilità** (20 fasce × N docenti con checkbox):
Usare `ttk.Treeview` in modalità "checkbutton" simulata: cella = "X" o "".
Singolo click alterna il valore e aggiorna il DataFrame in memoria.

---

## Decision 3: PyInstaller + OR-Tools packaging

**Decision**: `--onedir` (non `--onefile`) con `collect_all('ortools')` nel
file `.spec` di PyInstaller.

**Rationale**: OR-Tools 9.x include binari nativi (.so/.dll/.dylib). Con
`--onedir` i binari vengono copiati una volta nella cartella di distribuzione e
il tempo di avvio è minimo. Con `--onefile` ogni avvio estrae i binari in una
cartella temporanea (~5-10 s su Windows), violando SC-006 (avvio < 10 s).

**Alternatives considered**:
- `--onefile`: più comodo per la distribuzione ma avvio lento
- Escludere OR-Tools e usare solver pure-Python (es. python-constraint): nessun
  solver pure-Python regge il problema a questa scala con timeout 60 s

**Snippet .spec aggiuntivo**:
```python
from PyInstaller.utils.hooks import collect_all
datas_o, binaries_o, hiddenimports_o = collect_all('ortools')
a = Analysis(
    ['main.py'],
    ...
    datas=datas_o,
    binaries=binaries_o,
    hiddenimports=hiddenimports_o,
    ...
)
```

---

## Decision 4: reportlab PDF A4 orizzontale

**Decision**: `reportlab.platypus.Table` con `landscape(A4)` e larghezza
colonne calcolata dinamicamente in base al numero di docenti.

**Rationale**: Il numero di colonne (docenti) è variabile (10–30). Distribuire
la larghezza disponibile equamente garantisce che la tabella entri nel foglio.
Righe fisse: 20 fasce orarie + 1 intestazione = 21 righe totali.

**Layout**:
```
A4 orizzontale: 841.89 pt larghezza × 595.28 pt altezza
Margini: 20 pt ciascun lato
Larghezza disponibile: 801.89 pt
Prima colonna (etichetta fascia): 70 pt fissa
Colonne docenti: (801.89 - 70) / N_docenti pt ciascuna
```

**Alternanza colori righe** per giornata: ogni gruppo di 4 righe (stesso
giorno) ha sfondo alternato bianco/grigio chiaro per facilitare la lettura.

---

## Decision 5: Ordine di implementazione per user story

**Decision**: P2 (dati) → P1 (genera orario) → P3 (duplica anno)

**Rationale**: P2 è prerequisito di P1 (servono dati per generare). P3 riusa
interamente la logica di lettura/scrittura CSV di P2, quindi va implementato
dopo. Questa sequenza consente di validare i dati con lo scenario 1 di
quickstart.md prima di toccare il solver, e di avere un MVP (P1 + P2) prima
di aggiungere la funzione di duplicazione.

---

## Decision 6: Separatore CSV e codifica

**Decision**: Punto e virgola (`;`) come separatore, UTF-8 senza BOM come
codifica, per tutti i file CSV in input e output.

**Rationale**: Conforme all'assunzione nella spec. Excel su Windows con locale
italiano apre correttamente i CSV con `;` senza configurazione aggiuntiva.
UTF-8 senza BOM è compatibile sia con macOS che Windows (Python 3 default).

---

## Decision 7: Scrittura atomica dei CSV

**Decision**: Scrivere in un file temporaneo nella stessa cartella, poi
rinominare con `os.replace()` (atomico su POSIX e Windows NTFS).

**Rationale**: Rispetta il principio II (Data Integrity) della constitution.
Se la scrittura fallisce a metà, il file originale rimane intatto. `os.replace()`
è atomico sullo stesso filesystem su entrambe le piattaforme target.

```python
import os, tempfile
def scrivi_csv_atomico(df, path):
    dir_ = os.path.dirname(path)
    with tempfile.NamedTemporaryFile('w', dir=dir_, delete=False,
                                     suffix='.tmp', encoding='utf-8') as f:
        df.to_csv(f, sep=';', index=False)
        tmp = f.name
    os.replace(tmp, path)
```
