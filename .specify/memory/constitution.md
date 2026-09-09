<!--
SYNC IMPACT REPORT
==================
Version change: 1.0.0 → 2.0.0 (MAJOR: NON-NEGOTIABLE TDD mandate removed;
  Interoperability principle replaced by Desktop-First & Offline-Only)

Modified principles:
  - III. Test-First Development (NON-NEGOTIABLE) → REMOVED
  - IV. Interoperability → replaced by III. Desktop-First & Offline-Only (new)
  - New: IV. Standard Output Formats Only (added)
  - V. Simplicity → unchanged, renumbered (was already V)

Added sections:
  - III. Desktop-First & Offline-Only (new principle)
  - IV. Standard Output Formats Only (new principle)
  - Technology Stack (TODO resolved — full stack now defined)

Removed sections:
  - Principle III: Test-First Development (TDD mandate, NON-NEGOTIABLE label)
  - Principle IV: Interoperability (external calendar sync, iCal/ICS, Google Calendar)

Templates reviewed:
  - .specify/templates/plan-template.md   ✅ No changes required (Testing field is
      advisory; Constitution Check is generic; tests/ in structure is just an option)
  - .specify/templates/spec-template.md   ✅ No changes required (User Scenarios
      section covers acceptance criteria, not TDD workflow)
  - .specify/templates/tasks-template.md  ✅ No changes required (tests already
      marked OPTIONAL with explicit warning not to include unless requested)

Follow-up TODOs: None — all prior TODOs resolved.
-->

# Schedule Planner Constitution

## Core Principles

### I. User-Centric Design

Every feature MUST be grounded in a concrete user story with clearly stated
acceptance criteria. UI and workflow decisions MUST be justified by user value,
not technical convenience. Features without a traceable user story MUST NOT
be implemented.

**Rationale**: A schedule planner exists to serve users managing time and
commitments. Complexity added for its own sake directly harms the product.

### II. Data Integrity

Schedule and planning data MUST never be silently lost or corrupted. All write
operations to local CSV and PDF output files MUST complete fully before
reporting success. Any failure MUST surface a clear, actionable error to the
user — silent swallowing of file write errors is prohibited.

**Rationale**: Users depend on exported schedules for real-world coordination.
A silently incomplete PDF or corrupt CSV destroys trust and may cause planning
failures with no recovery path.

### III. Desktop-First & Offline-Only

The application MUST run entirely offline with no network requests, no web
framework, no remote API calls, and no database engine. All data MUST reside
as local files on the user's machine. The GUI MUST be implemented using
customtkinter. Web-based UI, REST APIs, cloud sync, and any networked
dependency are prohibited.

**Rationale**: The app is a portable local tool targeting Windows and macOS.
Introducing web or database dependencies breaks the portability guarantee,
creates unnecessary infrastructure complexity, and undermines the offline
design contract.

### IV. Standard Output Formats Only

The application MUST produce output exclusively as PDF (via reportlab) and/or
CSV (via pandas). No other output channels — databases, REST endpoints,
WebSockets, iCal/ICS, or proprietary binary formats — are permitted.
Imports, if supported, MUST read from CSV files only.

**Rationale**: PDF and CSV are universally readable without special tools,
require no runtime dependencies to open, and align with the offline-portable
mandate. Scope creep into additional formats adds maintenance burden without
user benefit.

### V. Simplicity

Every abstraction, dependency, or architectural layer MUST be justified by a
concrete, current need. YAGNI applies: no design-for-hypothetical-future-
requirements. Three similar lines of code are preferable to a premature
abstraction. Complexity violations MUST be documented in the plan's
Complexity Tracking table.

**Rationale**: The app's core logic (conflict detection, schedule optimization
via OR-Tools, PDF layout via reportlab) is already non-trivial. Unnecessary
structural complexity makes the codebase harder to maintain and harder to
package reliably with PyInstaller.

## Technology Stack

- **Language**: Python 3.11+
- **GUI**: customtkinter
- **Data manipulation**: pandas
- **PDF generation**: reportlab
- **Schedule optimization**: OR-Tools (Google)
- **Persistence**: local CSV files (no database of any kind)
- **Distribution**: PyInstaller — portable executable for macOS and Windows
- **Target platforms**: macOS and Windows (offline, no installer beyond the executable)

No additional dependency may be introduced without a constitution amendment
that explicitly justifies the addition against Principle V (Simplicity) and
Principle III (Desktop-First & Offline-Only).

## Development Workflow

All code changes MUST go through a pull request reviewed by at least one
other contributor before merging to the main branch.

Quality gates that MUST pass before merge:
- Manual verification of the acceptance criteria defined in the feature spec
- No TODOs or placeholder tokens introduced without a linked follow-up task
- Constitution Check in plan.md acknowledged and signed off
- PyInstaller build passes on at least one target platform before any release

Releases MUST be tagged with semantic version numbers. Breaking changes to
user-facing CSV schemas or PDF layout contracts require a MAJOR version bump
and a migration guide.

## Governance

This constitution supersedes all other development practices and conventions.
When a practice conflicts with a principle here, the constitution governs.

**Amendment procedure**: Any principle change MUST be documented as a
constitution update (via `/speckit-constitution`) before the changed practice
is applied. Amendments require:
1. A clear statement of what changes and why.
2. A version bump following semantic versioning rules.
3. A propagation check against all dependent templates.

**Versioning policy**:
- MAJOR: Backward-incompatible governance changes, principle removals, or
  redefinitions that invalidate prior decisions.
- MINOR: New principle or section added, or materially expanded guidance.
- PATCH: Clarifications, wording fixes, non-semantic refinements.

**Compliance review**: The Constitution Check section in each `plan.md` MUST
be completed before Phase 0 research begins and re-checked after Phase 1
design. Complexity violations MUST be justified in the plan's Complexity
Tracking table.

**Version**: 2.0.0 | **Ratified**: 2026-06-28 | **Last Amended**: 2026-06-29
