# Schema Griglia Output

La `GrigliaOrario` è la struttura condivisa tra `solver/scheduler.py`,
`output/pdf_export.py` e `output/csv_export.py`.

## Struttura Python

```python
# dict[tuple[int, int], dict[str, str]]
griglia: dict[tuple[int, int], dict[str, str]]
```

- **Chiave esterna**: `(giorno, fascia)` — giorno ∈ [0,4], fascia ∈ [0,3]
- **Chiave interna**: nome del docente (stringa; corrisponde a
  `DisponibilitaDocente.docente`)
- **Valore**: `"COGNOME (cl.X)"` se c'è una lezione assegnata; `""` se la
  fascia è libera per quel docente

**Esempio**:
```python
griglia = {
    (0, 0): {"Bianchi": "Rossi (cl.3)", "Ferrari": "", "Verdi": ""},
    (0, 1): {"Bianchi": "Conti (cl.1)", "Ferrari": "Neri (cl.5)", "Verdi": "Rossi (cl.3)"},
    # ... 20 coppie (giorno, fascia) totali
}
```

## Mapping giorni e fasce

| Indice | Giorno | Etichetta riga |
|--------|--------|----------------|
| 0 | Lunedì | `"Lunedì 13:30"` … `"Lunedì 16:30"` |
| 1 | Martedì | `"Martedì 13:30"` … |
| 2 | Mercoledì | `"Mercoledì 13:30"` … |
| 3 | Giovedì | `"Giovedì 13:30"` … |
| 4 | Venerdì | `"Venerdì 13:30"` … |

| Indice | Orario | Label |
|--------|--------|-------|
| 0 | 13:30–14:30 | `"13:30"` |
| 1 | 14:30–15:30 | `"14:30"` |
| 2 | 15:30–16:30 | `"15:30"` |
| 3 | 16:30–17:30 | `"16:30"` |

**Etichetta riga completa**: `f"{GIORNI[g]} {FASCE_LABEL[f]}"` — es. `"Lunedì 13:30"`

## Output CSV (`orario_generato.csv`)

Struttura del file esportato:

- **Riga 1 (intestazione)**: `Fascia;Docente1;Docente2;...` (docenti in ordine
  alfabetico)
- **Righe 2–21**: una per ogni (giorno, fascia) in ordine lun→ven, 13:30→16:30

```
Fascia;Bianchi;Ferrari;Verdi
Lunedì 13:30;Rossi (cl.3);;
Lunedì 14:30;Conti (cl.1);Neri (cl.5);Rossi (cl.3)
Lunedì 15:30;;;
Lunedì 16:30;;Bruni (cl.3);
Martedì 13:30;Galli (cl.4);;
...
```

Separatore: `;` | Codifica: UTF-8 | Nessuna riga indice pandas

## Layout PDF

| Elemento | Specifica |
|----------|-----------|
| Formato pagina | A4 orizzontale (`landscape(A4)`) |
| Margini | 20 pt tutti i lati |
| Prima colonna | "Fascia" — larghezza fissa 70 pt |
| Colonne docenti | Larghezza uguale = (801.89 − 70) / N_docenti pt |
| Righe | 1 intestazione + 20 dati = 21 righe |
| Contenuto cella | Cognome + classe: `"Rossi (cl.3)"` |
| Cella vuota | Bianca, nessun testo |
| Intestazione col. | Bold, sfondo grigio chiaro |
| Raggruppamento | Ogni gruppo di 4 righe (stesso giorno) ha sfondo alternato bianco / grigio molto chiaro per facilitare la lettura |
| Font | Helvetica 8pt per le celle dati; Helvetica-Bold 9pt per le intestazioni |
