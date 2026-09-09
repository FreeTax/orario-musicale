# Implementation Plan: Orario Pomeridiano Liceo Musicale

**Branch**: `001-orario-musicale` | **Date**: 2026-06-29 | **Spec**: [spec.md](spec.md)

**Input**: Feature specification from `specs/001-orario-musicale/spec.md`

## Summary

App desktop offline per automatizzare la costruzione dell'orario pomeridiano
di un liceo musicale. Il coordinatore importa tre CSV (studenti, gruppi LMC,
disponibilità docenti), avvia il solver CP-SAT di OR-Tools che rispetta i
vincoli hard (disponibilità docenti, no doppia prenotazione, co-scheduling LMC)
e ottimizza i vincoli soft (buchi, rientri, distanza, ore consecutive, giorno
unico), poi esporta la griglia risultante in PDF A4 orizzontale e CSV.

**Ordine implementativo**: P2 (gestione dati) → P1 (genera orario + export) →
P3 (duplica anno con promozione).

## Technical Context

**Language/Version**: Python 3.13

**Primary Dependencies**: customtkinter 5.x, pandas 2.x, reportlab 4.x,
OR-Tools 9.x (CP-SAT), PyInstaller 6.x

**Storage**: File CSV locali in `OrarioMusicale/<anno>/`; nessun database

**Testing**: Verifica manuale dei criteri di accettazione — vedere
`quickstart.md` per gli scenari di validazione (no TDD per constitution v2.0.0)

**Target Platform**: macOS e Windows (desktop app, completamente offline)

**Project Type**: desktop-app

**Performance Goals**: Generazione orario < 2 minuti; timeout solver 60 s;
avvio app < 10 s; PDF stampabile su A4 orizzontale

**Constraints**: Completamente offline; nessuna rete; portatile via PyInstaller
(`.exe` Windows, `.app` macOS); separatore CSV `;`, codifica UTF-8

**Scale/Scope**: ~100–200 studenti/anno; ~20–30 docenti; 20 fasce orarie
(5 giorni × 4 fasce); griglia output ~20 righe × N colonne docenti

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principio | Stato | Note |
|-----------|-------|------|
| I. User-Centric Design | ✅ PASS | Ogni schermata è tracciata a una user story nella spec; nessuna funzionalità senza storia utente |
| II. Data Integrity | ✅ PASS | CSV scritti atomicamente (write→rename); solver non restituisce output parziali su errore hard; errori mostrati esplicitamente all'utente |
| III. Desktop-First & Offline-Only | ✅ PASS | customtkinter GUI; nessuna rete; nessun DB; tutti i dati in file CSV locali |
| IV. Standard Output Formats Only | ✅ PASS | Solo PDF via reportlab e CSV via pandas; nessun altro canale di output |
| V. Simplicity | ✅ PASS | 4 moduli flat (gui/, solver/, data/, output/); nessun ORM, nessun framework web, nessuna astrazione prematura; solver invocabile direttamente senza GUI |

**Post-design re-check** (dopo Phase 1): ✅ La struttura a moduli non introduce
layer non necessari. `data/` è puro I/O CSV; `solver/` è isolato dalla GUI;
`output/` è puramente funzionale. Nessuna violazione da giustificare.

## Project Structure

### Documentation (this feature)

```text
specs/001-orario-musicale/
├── plan.md              # Questo file
├── research.md          # Decisioni tecniche Phase 0
├── data-model.md        # Entità e strutture dati Phase 1
├── quickstart.md        # Guida di validazione Phase 1
├── contracts/           # Schemi CSV, interfaccia solver, griglia output
│   ├── csv-schemas.md
│   ├── solver-interface.md
│   └── output-grid.md
└── tasks.md             # Phase 2 output (/speckit-tasks — non creato da /speckit-plan)
```

### Source Code (repository root)

```text
schedulePlanner/
├── main.py                      # Entry point: inizializza CTk app, carica Home
├── gui/
│   ├── home.py                  # Schermata 1: lista anni, Nuovo/Apri/Duplica
│   ├── studenti.py              # Schermata 2: tabella studenti, import/export
│   ├── gruppi_lmc.py            # Schermata 3: tabella gruppi LMC, import/export
│   ├── disponibilita.py         # Schermata 4: griglia disponibilità, import/export
│   ├── genera_orario.py         # Schermata 5: genera, anteprima, export PDF/CSV
│   └── duplica_anno.py          # Schermata 6: wizard promozione studenti
├── solver/
│   └── scheduler.py             # OR-Tools CP-SAT: dati → griglia orario
├── data/
│   ├── studenti.py              # Legge/scrive studenti.csv
│   ├── gruppi_lmc.py            # Legge/scrive gruppi_lmc.csv
│   └── disponibilita.py         # Legge/scrive disponibilita_docenti.csv
├── output/
│   ├── pdf_export.py            # Genera PDF A4 orizzontale da griglia
│   └── csv_export.py            # Genera CSV da griglia
└── OrarioMusicale/              # Dati runtime (non nel repository)
    └── 2026-27/
        ├── studenti.csv
        ├── gruppi_lmc.csv
        ├── disponibilita_docenti.csv
        └── orario_generato.csv  # Output solver (se generato)
```

**Structure Decision**: Single desktop project. Nessuna separazione
frontend/backend (tutto locale). La separazione `gui/` ↔ `solver/` ↔ `data/`
↔ `output/` garantisce che il solver sia invocabile indipendentemente dalla GUI
e che ogni modulo abbia una singola responsabilità.

## Complexity Tracking

> Nessuna violazione dei principi della constitution. Tabella non necessaria.
