# Data Model: Orario Pomeridiano Liceo Musicale

**Feature**: `001-orario-musicale`
**Date**: 2026-06-29

## Entità

### Studente

Rappresenta uno studente iscritto per un anno scolastico.

| Campo | Tipo Python | Validazione |
|-------|-------------|-------------|
| cognome | str | Non vuoto |
| nome | str | Non vuoto |
| classe | int | 1–5 |
| km | float | ≥ 0; default 0.0 se campo assente o non numerico |
| docente_1str | str | Non vuoto |
| strumento_1 | str | Non vuoto |
| docente_2str | str | Vuoto se classe 5; non vuoto altrimenti |
| strumento_2 | str | Vuoto se classe 5; non vuoto altrimenti |
| ore_consecutive | str | "SI" o "NO" (normalizzato upper-strip) |
| giorno_unico | str | "SI" o "NO" (normalizzato upper-strip) |

**Proprietà derivate** (calcolate, non memorizzate nel CSV):

| Proprietà | Formula |
|-----------|---------|
| `ha_2str` | `classe in {1,2,3,4}` AND `docente_2str != ""` |
| `ha_lmc` | `classe in {3,4,5}` |
| `ore_1str` | `2` se `classe in {1,2,5}`; `1` se `classe in {3,4}` |
| `limite_giorni` | `3` se `km < 5`; `2` altrimenti |

**Identificatore**: `cognome` — si assume unicità all'interno di un anno scolastico

---

### GruppoLMC

Gruppo di musica da camera con un docente e da 2 a 4 studenti.

| Campo | Tipo Python | Validazione |
|-------|-------------|-------------|
| numero_gruppo | int | > 0, univoco per anno |
| docente | str | Non vuoto; deve esistere in DisponibilitaDocente |
| studenti | list[str] | 2–4 cognomi; ogni cognome deve esistere in Studente; solo classi 3–5 |

**Struttura CSV**: colonne fisse `studente_1`, `studente_2`, `studente_3`,
`studente_4`; le colonne vuote vengono scartate in lettura.

---

### DisponibilitaDocente

Disponibilità settimanale di un docente su 20 fasce orarie.

| Campo | Tipo Python | Validazione |
|-------|-------------|-------------|
| docente | str | Non vuoto, univoco |
| disponibilita | list[list[bool]] | 5 giorni × 4 fasce; True se CSV contiene "X" (case-insensitive) |

**Accesso**: `disponibilita[giorno][fascia]` dove giorno ∈ [0,4] e fascia ∈ [0,3]

**Mapping colonne CSV → indici**:

| Indice giorno | Giorno | Colonne CSV |
|---------------|--------|-------------|
| 0 | Lunedì | lun_1330, lun_1430, lun_1530, lun_1630 |
| 1 | Martedì | mar_1330, mar_1430, mar_1530, mar_1630 |
| 2 | Mercoledì | mer_1330, mer_1430, mer_1530, mer_1630 |
| 3 | Giovedì | gio_1330, gio_1430, gio_1530, gio_1630 |
| 4 | Venerdì | ven_1330, ven_1430, ven_1530, ven_1630 |

| Indice fascia | Orario |
|---------------|--------|
| 0 | 13:30–14:30 |
| 1 | 14:30–15:30 |
| 2 | 15:30–16:30 |
| 3 | 16:30–17:30 |

---

### AnnoScolastico

Metadato per un anno scolastico. Non ha un file dedicato: è la cartella stessa.

| Campo | Tipo Python | Validazione |
|-------|-------------|-------------|
| label | str | Formato "AAAA-AA" (es. "2026-27") |
| path | pathlib.Path | `OrarioMusicale/<label>/`; deve esistere su disco |

**Presenza file**: la cartella è valida se esiste; file CSV assenti sono
trattati come tabelle vuote (nessun errore di apertura anno).

---

### Lezione (entità interna al solver)

Unità atomica di scheduling: un'ora da assegnare a una fascia.

| Campo | Tipo Python | Note |
|-------|-------------|------|
| id | str | Chiave univoca: `f"{cognome}_{tipo}_{indice}"` |
| cognome_studente | str | Riferimento a Studente.cognome |
| docente | str | Docente responsabile della lezione |
| tipo | str | `"1str"` \| `"2str"` \| `"lmc"` |
| indice | int | 1 o 2 (per le 2 ore di 1°str); 1 altrimenti |
| gruppo_lmc | int \| None | `numero_gruppo` se tipo="lmc"; None altrimenti |

**Generazione**: il layer `data/` costruisce la lista di `Lezione` a partire
dagli oggetti `Studente` e `GruppoLMC` prima di passarla al solver.

---

### Assegnazione (output solver per una lezione)

| Campo | Tipo Python | Note |
|-------|-------------|------|
| lezione_id | str | Riferimento a Lezione.id |
| giorno | int | 0=lunedì … 4=venerdì |
| fascia | int | 0=13:30, 1=14:30, 2=15:30, 3=16:30 |

---

### GrigliaOrario (struttura di output condivisa)

Struttura intermedia usata da `pdf_export.py` e `csv_export.py`.

```python
# Tipo: dict[tuple[int, int], dict[str, str]]
griglia: dict[tuple[int, int], dict[str, str]]
# griglia[(giorno, fascia)][docente] = "COGNOME (cl.X)" | ""
```

**Costruzione**: dalla lista di `Assegnazione` + mappa `lezione_id → Studente`.

---

## Relazioni

```
AnnoScolastico
  ├── lista[Studente]
  ├── lista[GruppoLMC]
  └── lista[DisponibilitaDocente]

GruppoLMC.studenti           ──ref──> Studente.cognome
GruppoLMC.docente            ──ref──> DisponibilitaDocente.docente
Studente.docente_1str        ──ref──> DisponibilitaDocente.docente
Studente.docente_2str        ──ref──> DisponibilitaDocente.docente (se non vuoto)

Solver pipeline:
  lista[Studente] + lista[GruppoLMC] + dict[DisponibilitaDocente]
    → lista[Lezione]             (costruita dal layer data/)
    → lista[Assegnazione]        (output OR-Tools CP-SAT)
    → GrigliaOrario              (costruita post-solve)
    → PDF / CSV                  (output/)
```

## Validazione cross-entità (pre-solver)

Questi controlli DEVONO essere eseguiti da `data/` prima di passare i dati
al solver, e DEVONO produrre un errore leggibile se falliscono:

1. Ogni `GruppoLMC.docente` esiste in `DisponibilitaDocente`
2. Ogni `GruppoLMC.studenti[i]` esiste in `Studente.cognome`
3. Ogni `Studente.docente_1str` esiste in `DisponibilitaDocente`
4. Ogni `Studente.docente_2str` (se non vuoto) esiste in `DisponibilitaDocente`
5. Studenti nei gruppi LMC appartengono a classi 3, 4 o 5
6. I gruppi LMC hanno da 2 a 4 studenti
