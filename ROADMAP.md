# Roadmap

## Phase 1 - Snapshot Import

- Imported the latest ProjectLibre SourceForge snapshot into `upstream/projectlibre-snapshot`.
- Captured provenance in `NOTICE` and `upstream/README.md`.
- Create initial snapshot archive placeholder if the real snapshot cannot be fetched immediately.
- Add CPAL license notice in `upstream/NOTICE` and ensure LICENSEs accompany the snapshot.

## Phase 2 - Domain Model

- Define Rust types for projects, tasks, dependencies, resources, assignments, calendars, and baselines.
- Add validation, serialization, and schedule recalculation for the workspace state.
- Create a standalone Rust crate under `projectlibre-tauri/backend` for domain models and core logic.

## Phase 3 - Persistence and Commands

- Store workspace state locally in JSON first.
- Move the core workspace model and store into the Rust backend crate.
- Add XML interchange so the Rust backend can import/export workspace state in a portable format.
- Expose Tauri commands for project/task/dependency/resource/assignment/calendar CRUD.
- Add export/import, edit, and schedule-rebuild commands.
- Add tests for delete cascades, load/save round trips, XML round trips, and schedule ordering.
- Implement a basic persistence layer in the backend and wire with simple tests.

## Phase 4 - Frontend

- Replace placeholder UI with project navigation and task editing screens.
- Feed a basic Gantt-style timeline from the Rust state.
- Integrate frontend with backend via TAURI commands and expose initial UI components.

## Phase 5 - Feature Growth

- Add calendar rules, task sorting, and import/export.
- Expand coverage in small slices instead of a large rewrite.
- Establish a polite feature roadmap for future iterations (Clustering tasks, resource leveling, and cloud sync considerations).
