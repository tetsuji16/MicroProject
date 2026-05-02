# Roadmap

## Phase 1 - Snapshot Import

- Imported the latest ProjectLibre SourceForge snapshot into `upstream/projectlibre-snapshot`.
- Captured provenance in `NOTICE` and `upstream/README.md`.

## Phase 2 - Domain Model

- Define Rust types for projects, tasks, dependencies, and calendars.
- Add validation and serialization for the workspace state.

## Phase 3 - Persistence and Commands

- Store workspace state locally in JSON first.
- Expose Tauri commands for project/task/dependency CRUD.
- Add tests for delete cascades and load/save round trips.

## Phase 4 - Frontend

- Replace placeholder UI with project navigation and task editing screens.
- Feed a basic Gantt-style timeline from the Rust state.

## Phase 5 - Feature Growth

- Add calendar rules, task sorting, and import/export.
- Expand coverage in small slices instead of a large rewrite.

