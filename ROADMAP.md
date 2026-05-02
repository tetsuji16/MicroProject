# Roadmap

## Phase 1 - Snapshot Import

- Imported the latest ProjectLibre SourceForge snapshot into `upstream/projectlibre-snapshot`.
- Captured provenance in `NOTICE` and `upstream/README.md`.

## Phase 2 - Domain Model

- Define Rust types for projects, tasks, dependencies, resources, assignments, calendars, and baselines.
- Add validation, serialization, and schedule recalculation for the workspace state.

## Phase 3 - Persistence and Commands

- Store workspace state locally in JSON first.
- Add XML interchange so the Rust backend can import/export workspace state in a portable format.
- Expose Tauri commands for project/task/dependency/resource/assignment/calendar CRUD.
- Add export/import and schedule-rebuild commands.
- Add tests for delete cascades, load/save round trips, and schedule ordering.

## Phase 4 - Frontend

- Replace placeholder UI with project navigation and task editing screens.
- Feed a basic Gantt-style timeline from the Rust state.

## Phase 5 - Feature Growth

- Add calendar rules, task sorting, and import/export.
- Expand coverage in small slices instead of a large rewrite.
