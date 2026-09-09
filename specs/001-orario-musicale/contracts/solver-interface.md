# Interfaccia Solver — `solver/scheduler.py`

## Firma della funzione principale

```python
# solver/scheduler.py

from dataclasses import dataclass, field

GIORNI = ["lun", "mar", "mer", "gio", "ven"]
FASCE  = [1330, 1430, 1530, 1630]


@dataclass
class RisultatoOK:
    griglia: dict          # GrigliaOrario: {(giorno, fascia): {docente: "COGNOME (cl.X)"}}
    violazioni_soft: list  # list[str] — descrizione soft non soddisfatti; vuota se ottimale


@dataclass
class RisultatoErrore:
    motivo: str            # Descrizione del vincolo hard non soddisfacibile


def genera_orario(
    studenti: list[dict],
    gruppi_lmc: list[dict],
    disponibilita: dict,
    timeout_sec: int = 60,
) -> RisultatoOK | RisultatoErrore:
    ...
```

## Tipi degli argomenti

### `studenti: list[dict]`

Ogni dizionario corrisponde a una riga di `studenti.csv` dopo il parsing:

```python
{
    "cognome": str,
    "nome": str,
    "classe": int,            # 1–5
    "km": float,
    "docente_1str": str,
    "strumento_1": str,
    "docente_2str": str,      # "" per classe 5
    "strumento_2": str,       # "" per classe 5
    "ore_consecutive": str,   # "SI" o "NO"
    "giorno_unico": str,      # "SI" o "NO"
}
```

### `gruppi_lmc: list[dict]`

```python
{
    "numero_gruppo": int,
    "docente": str,
    "studenti": list[str],    # 2–4 cognomi
}
```

### `disponibilita: dict[str, list[list[bool]]]`

```python
# disponibilita[nome_docente][giorno_idx][fascia_idx] = True | False
# giorno_idx: 0=lun … 4=ven
# fascia_idx: 0=13:30 … 3=16:30
```

## Contratto comportamentale

| Condizione | Tipo restituito | Contenuto |
|-----------|-----------------|-----------|
| Tutti i vincoli hard soddisfatti, soft ottimali | `RisultatoOK` | `griglia` completa, `violazioni_soft = []` |
| Tutti i vincoli hard soddisfatti, soft parziali (timeout 60s) | `RisultatoOK` | `griglia` completa, `violazioni_soft` descrive i soft non soddisfatti |
| Almeno un vincolo hard non soddisfacibile | `RisultatoErrore` | `motivo` descrive il conflitto (es. "Docente Bianchi: disponibilità insufficiente per 3 lezioni") |

**Garanzie**:
- Il solver non lancia eccezioni per dati validi; le eccezioni indicano bug
- Mai restituisce una griglia parziale: o è completa o è `RisultatoErrore`
- I vincoli hard non vengono mai violati nell'output `RisultatoOK`

## Precondizioni (validate da `data/` prima di invocare il solver)

Il solver presuppone che i dati siano già stati validati:
- Tutti i docenti in `studenti` e `gruppi_lmc` esistono in `disponibilita`
- Tutti i cognomi in `gruppi_lmc.studenti` esistono in `studenti`
- I membri dei gruppi LMC appartengono a classi 3–5
- I campi booleani (`ore_consecutive`, `giorno_unico`) sono "SI" o "NO"
- `classe` è un intero in [1, 5]

Passare dati non validati produce comportamento indefinito.
