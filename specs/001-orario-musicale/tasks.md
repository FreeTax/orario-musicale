---
description: "Task list for Orario Pomeridiano Liceo Musicale"
---

# Tasks: Orario Pomeridiano Liceo Musicale

**Input**: Design documents from `specs/001-orario-musicale/`

**Prerequisites**: plan.md ✅ | spec.md ✅ | research.md ✅ | data-model.md ✅ | contracts/ ✅ | quickstart.md ✅

**Tests**: No test tasks — manual validation via quickstart.md (constitution v2.0.0 removes TDD mandate).

**Organization**: Tasks grouped by user story per implementation order (US2 → US1 → US3).

## Format: `[ID] [P?] [Story?] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: [US1] = Genera orario | [US2] = Importa dati | [US3] = Duplica anno

## Path Conventions

All source paths are relative to the repository root: `schedulePlanner/`

---

## Phase 1: Setup (Shared Infrastructure)

**Purpose**: Project initialization — directory structure, dependencies, app skeleton.

- [ ] T001 Create `schedulePlanner/` project directories: `gui/`, `solver/`, `data/`, `output/`, `OrarioMusicale/` with empty `__init__.py` files in each Python package
- [ ] T002 [P] Create `schedulePlanner/requirements.txt` listing: `customtkinter>=5.2`, `pandas>=2.0`, `reportlab>=4.0`, `ortools>=9.8`, `pyinstaller>=6.0`
- [ ] T003 Create `schedulePlanner/main.py` — CTkApp subclass with `CTkTabview` or frame-based navigation; registers the 6 screen frames (Home, Studenti, GruppiLMC, Disponibilita, GeneraOrario, DuplicaAnno); calls `app.mainloop()`

---

## Phase 2: Foundational (Blocking Prerequisites)

**Purpose**: Data I/O layer and year management — MUST be complete before any user story can be implemented.

**⚠️ CRITICAL**: No user story work can begin until this phase is complete.

- [ ] T004 [P] Implement `schedulePlanner/data/studenti.py` — `carica_studenti(path: Path) -> list[dict]` reads `studenti.csv` (sep=`;`, utf-8), normalises `ore_consecutive`/`giorno_unico` to uppercase, coerces `km` to float (default 0.0 on missing/invalid), raises `ValueError` with row number on missing mandatory fields; `salva_studenti(studenti: list[dict], path: Path)` writes atomically via `os.replace()`
- [ ] T005 [P] Implement `schedulePlanner/data/gruppi_lmc.py` — `carica_gruppi(path: Path) -> list[dict]` reads `gruppi_lmc.csv`, reconstructs `studenti: list[str]` from `studente_1..4` columns (skipping empty cells); `salva_gruppi(gruppi: list[dict], path: Path)` writes atomically
- [ ] T006 [P] Implement `schedulePlanner/data/disponibilita.py` — `carica_disponibilita(path: Path) -> dict[str, list[list[bool]]]` reads `disponibilita_docenti.csv`, converts "X" cells to `True` (case-insensitive) and empty to `False`, returns `disp[docente][giorno][fascia]` (5×4 matrix); `salva_disponibilita(disp: dict, path: Path)` writes atomically
- [ ] T007 Implement year management functions in `schedulePlanner/main.py` — `lista_anni(base: Path) -> list[str]` scans `OrarioMusicale/` for subdirectories; `crea_anno(base: Path, label: str) -> Path` creates the year folder; `apri_anno(base: Path, label: str) -> dict` returns `{"studenti": [...], "gruppi": [...], "disponibilita": {...}}` loading all three CSVs (empty list/dict if file absent)
- [ ] T008 Implement `schedulePlanner/gui/home.py` — `CTkScrollableFrame` listing year folders from `lista_anni()`; buttons: "Nuovo Anno" (dialog for label, calls `crea_anno()`), "Apri" (calls `apri_anno()`, switches to Studenti screen), "Duplica" (navigates to DuplicaAnno screen with selected year pre-filled)

**Checkpoint**: Foundation ready — all three user story phases can now be worked on.

---

## Phase 3: User Story 2 — Importa e gestisci dati (Priority: P2) 🗂️

**Goal**: Coordinator can import, edit, and export all three CSV types for a school year.

**Independent Test**: Open "test-validazione" → import all 3 CSVs → edit one record in each screen → export CSVs → close and reopen app → verify all changes persisted. (quickstart.md Scenario 1)

### Implementation for User Story 2

- [ ] T009 [P] [US2] Implement `schedulePlanner/gui/studenti.py` — `CTkFrame` with `ttk.Treeview` displaying all 10 student columns; double-click on cell spawns a `CTkEntry` overlay for inline editing (confirmed on Enter/FocusOut); "Importa CSV" button opens `filedialog.askopenfilename()`, calls `carica_studenti()`, refreshes view; "Esporta CSV" calls `salva_studenti()` with current table data; changes are held in memory and persisted to disk on explicit save or screen switch
- [ ] T010 [P] [US2] Implement `schedulePlanner/gui/gruppi_lmc.py` — `CTkFrame` with `ttk.Treeview` for 6 columns (`numero_gruppo`, `docente`, `studente_1..4`); same inline editing, Import/Export CSV pattern as T009 using `carica_gruppi()` / `salva_gruppi()`
- [ ] T011 [P] [US2] Implement `schedulePlanner/gui/disponibilita.py` — `CTkFrame` with `ttk.Treeview` showing `docente` column + 20 availability columns; single-click on an availability cell toggles "X" ↔ "" and updates the in-memory `dict[str, list[list[bool]]]`; Import/Export CSV via `carica_disponibilita()` / `salva_disponibilita()`

**Checkpoint**: User Story 2 fully functional — coordinator can manage all school year data independently.

---

## Phase 4: User Story 1 — Genera e scarica l'orario (Priority: P1) 🎯 MVP

**Goal**: Coordinator generates a valid schedule satisfying all hard constraints, previews it, and exports PDF and CSV.

**Independent Test**: With "test-validazione" data loaded, click "Genera orario" → verify grid satisfies all hard constraints from quickstart.md Scenario 2 → export PDF (A4 landscape, readable) → export CSV (structure with "Fascia" column + teacher columns).

### Implementation for User Story 1

- [ ] T012 [P] [US1] Implement `schedulePlanner/solver/scheduler.py` — define `RisultatoOK` and `RisultatoErrore` dataclasses; implement `genera_orario(studenti, gruppi_lmc, disponibilita, timeout_sec=60)`: (1) cross-entity validation (all teacher refs exist in disponibilita, all LMC student refs exist in studenti, LMC members are classes 3-5), returning `RisultatoErrore` on first violation; (2) build `Lezione` list from student class rules (data-model.md); (3) CP-SAT model with binary vars `x[lesson_id, day, slot]`, hard constraints H1–H5 per contracts/solver-interface.md, weighted soft objective (buchi×100, giorni_extra×50, distanza×20, ore_consecutive×80, giorno_unico×80); (4) set `solver.parameters.max_time_in_seconds = timeout_sec`; (5) return `RisultatoOK` with `GrigliaOrario` dict and soft violation descriptions, or `RisultatoErrore` if INFEASIBLE
- [ ] T013 [P] [US1] Implement `schedulePlanner/output/pdf_export.py` — `esporta_pdf(griglia: dict, docenti: list[str], path: Path)`: creates `landscape(A4)` PDF using `reportlab.platypus.Table`; first column "Fascia" 70pt wide; remaining columns split equally across available width; 20 data rows + header row; cell format `"COGNOME (cl.X)"` or blank; alternating background per day group (4 rows each); saves to `path`
- [ ] T014 [P] [US1] Implement `schedulePlanner/output/csv_export.py` — `esporta_csv(griglia: dict, docenti: list[str], path: Path)`: builds pandas DataFrame with "Fascia" column + one column per docente (alphabetical order); 20 rows (lun–ven × 4 fasce); writes with `sep=';'`, utf-8, no index, atomically via `os.replace()`
- [ ] T015 [US1] Implement `schedulePlanner/gui/genera_orario.py` — "Genera orario" button: disables itself, shows `CTkProgressBar` (indeterminate), runs `genera_orario()` in `threading.Thread` to keep GUI responsive; on `RisultatoOK`: shows read-only `ttk.Treeview` preview of `GrigliaOrario`, enables "Esporta PDF" (calls `esporta_pdf()`) and "Esporta CSV" (calls `esporta_csv()`) buttons; on `RisultatoErrore`: shows `CTkLabel` with error message in red, keeps Export buttons disabled; generation can be re-run after data changes

**Checkpoint**: User Story 1 fully functional — full MVP (data import + schedule generation + export) is now working end-to-end.

---

## Phase 5: User Story 3 — Duplica anno con promozione (Priority: P3) 📅

**Goal**: Coordinator creates next school year from current one with automatic student promotion.

**Independent Test**: Duplicate "test-validazione" to "test-validazione-2" → verify wizard summary shows correct promotion table → confirm → open new year → verify class promotions, no class-5 students, empty class-1, empty LMC groups, unchanged teacher availability. (quickstart.md Scenario 4)

### Implementation for User Story 3

- [ ] T016 [US3] Implement promotion algorithm as `_calcola_promozione(studenti: list[dict]) -> dict` in `scheduleplanner/gui/duplica_anno.py` — returns `{"promossi": [...], "rimossi": [...], "nuova_lista": [...]}` where promoted students have `classe += 1`, class-5 students are moved to "rimossi", class-1 is cleared; function is pure (no side effects, takes current student list, returns transformed list)
- [ ] T017 [US3] Complete `schedulePlanner/gui/duplica_anno.py` wizard — multi-step `CTkToplevel`: Step 1: `CTkComboBox` for source year + `CTkEntry` for new label; Step 2: `ttk.Treeview` summary table showing each student's `cognome`, old class → new class (or "Rimosso"); "Conferma" button checks if `OrarioMusicale/<new_label>/` already exists (if so: `CTkMessagebox` warning with Sovrascivi/Annulla); on confirm: calls `crea_anno()`, writes promoted `studenti.csv` via `salva_studenti()`, writes empty `gruppi_lmc.csv` via `salva_gruppi([])`, copies `disponibilita_docenti.csv` unchanged; dismisses wizard and refreshes Home year list

**Checkpoint**: All three user stories independently functional and testable.

---

## Phase 6: Polish & Cross-Cutting Concerns

**Purpose**: Packaging and final validation.

- [ ] T018 [P] Create `schedulePlanner/schedulePlanner_win.spec` — PyInstaller spec for Windows: `Analysis(['main.py'], ...)` with `collect_all('ortools')` for datas/binaries/hiddenimports; `EXE(name='OrarioMusicale', onedir=True, console=False)`
- [ ] T019 [P] Create `schedulePlanner/schedulePlanner_mac.spec` — PyInstaller spec for macOS: same `collect_all('ortools')` setup; `BUNDLE(name='OrarioMusicale.app', bundle_identifier='studio.kinoa.orariomusicale')`
- [ ] T020 Execute all 5 scenarios in `specs/001-orario-musicale/quickstart.md` — run each scenario top-to-bottom, verify every ✅ check passes; document any failures as GitHub issues before marking complete
- [ ] T021 [P] Pin exact installed versions in `schedulePlanner/requirements.txt` after T020 passes (run `pip freeze | grep -E 'customtkinter|pandas|reportlab|ortools|pyinstaller'`)

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: No dependencies — start immediately
- **Foundational (Phase 2)**: Depends on Phase 1 — **BLOCKS all user stories**
- **US2 (Phase 3)**: Depends on Foundational only — T009/T010/T011 fully parallel
- **US1 (Phase 4)**: Depends on Foundational + US2 complete (solver needs loadable data) — T012/T013/T014 parallel, T015 depends on all three
- **US3 (Phase 5)**: Depends on Foundational + US2 complete (reads/writes CSVs) — T016 then T017
- **Polish (Phase 6)**: Depends on all user stories complete

### User Story Dependencies

- **US2 (P2)**: Can start after Foundational — no dependency on US1 or US3
- **US1 (P1)**: Depends on US2 being complete (data loading must work to feed solver)
- **US3 (P3)**: Depends on Foundational; can start in parallel with US1 (both need the data layer but neither needs the other)

### Within Phase 4 (US1)

- T012, T013, T014 are all independent (different files) — launch in parallel
- T015 depends on T012 (imports `genera_orario`), T013 (imports `esporta_pdf`), T014 (imports `esporta_csv`)

### Parallel Opportunities

- T002, T003 can run in parallel (Phase 1 after T001)
- T004, T005, T006 are fully independent (Phase 2 foundational data layer)
- T009, T010, T011 are fully independent (Phase 3 GUI screens)
- T012, T013, T014 are fully independent (Phase 4 solver, pdf, csv)
- T018, T019, T021 are fully independent (Phase 6 polish)

---

## Parallel Example: Phase 4 (US1)

```bash
# Launch T012, T013, T014 simultaneously (all different files, no dependencies):
Task: "Implement solver/scheduler.py — CP-SAT model + validation"
Task: "Implement output/pdf_export.py — landscape A4 PDF from GrigliaOrario"
Task: "Implement output/csv_export.py — CSV from GrigliaOrario"

# Once all three complete, launch T015:
Task: "Implement gui/genera_orario.py — integrates solver + pdf + csv"
```

---

## Implementation Strategy

### MVP First (US2 + US1 only)

1. Complete Phase 1: Setup
2. Complete Phase 2: Foundational (CRITICAL — blocks everything)
3. Complete Phase 3: User Story 2 (data management)
4. Complete Phase 4: User Story 1 (generate + export)
5. **STOP and VALIDATE**: Run quickstart.md Scenarios 1, 2, 3
6. Demo/deliver: full scheduling workflow is functional

### Incremental Delivery

1. Setup + Foundational → project skeleton + data layer ready
2. Add US2 → coordinator can manage school year data (CSV import/edit/export)
3. Add US1 → coordinator can generate, preview, and export schedule (full MVP!)
4. Add US3 → coordinator gains multi-year management (year duplication)
5. Polish → packaged .exe and .app ready for distribution

### Parallel Team Strategy (2 developers)

1. Both complete Phase 1 + Phase 2 together
2. Once Foundational is done:
   - Developer A: US2 (Phase 3: T009, T010, T011)
   - Developer B: Start US1 T012 + T013 + T014 in parallel
3. Developer A completes US2 → joins Developer B on T015 (gui/genera_orario.py)
4. Together: US3 (T016, T017) and Polish (T018–T021)

---

## Notes

- `[P]` = different files, no shared state dependencies — safe to run concurrently
- `[US1/2/3]` = maps task to spec.md user story for traceability
- Each user story phase is independently completable and testable via quickstart.md
- No test tasks (manual validation via quickstart.md per constitution v2.0.0)
- Commit after each task or logical group; each phase checkpoint is a good commit point
- Atomic CSV writes (`os.replace()`) required for all `salva_*` functions — see research.md Decision 7
- `ttk.Treeview` styling will not match CTk theme exactly — acceptable for enterprise internal app
