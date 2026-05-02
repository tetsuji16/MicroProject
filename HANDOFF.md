# MicroProject Handoff

This repository is the shared workspace for two connected tracks:

1. A latest-snapshot import of ProjectLibre from SourceForge, kept under `upstream/projectlibre-snapshot`.
2. A Rust + Tauri rewrite scaffold in `projectlibre-tauri`.

## Current Truth

- Repo purpose: `MicroProject` is the public home for the rewrite.
- Upstream provenance:
  - SourceForge tree: https://sourceforge.net/p/projectlibre/code/ci/master/tree/
  - Imported commit: `0530be227f4a10c5545cce8d3db20ac5a4d76a66`
- Licensing:
  - Upstream code remains CPAL 1.0 with its own notices inside the imported tree.
  - New rewrite work is documented separately at the repo root.

## Important Paths

- `README.md`: primary project overview and entry point.
- `NOTICE`: provenance and license summary.
- `ROADMAP.md`: phased implementation plan.
- `CONTRIBUTING.md`: how to contribute safely.
- `upstream/projectlibre-snapshot/`: imported ProjectLibre snapshot, treat as read-only unless refreshing upstream.
- `projectlibre-tauri/`: Rust + Tauri app scaffold.

## What Is Already Done

- Added root documentation and licensing files.
- Imported the ProjectLibre latest snapshot into `upstream/projectlibre-snapshot`.
- Implemented a working Tauri scaffold:
  - JSON-backed local workspace store
  - Rust models for projects, tasks, dependencies, calendars
  - Tauri commands for CRUD operations
  - Minimal frontend placeholder
- `cargo check` and `cargo test` pass in `projectlibre-tauri`.

## Next Best Steps

1. Add unit tests for `projectlibre-tauri/src/storage.rs`.
2. Add a small frontend that calls the Tauri commands.
3. Split out the first real UI slice for project and task management.
4. If upstream needs to be refreshed, replace the imported snapshot and update provenance notes.

## Notes For Future AI

- Do not collapse the upstream snapshot into the rewrite code. Keep provenance visible.
- Keep `main` focused on the rewrite, not on full upstream history mirroring.
- Prefer incremental, testable vertical slices over a big-bang port.
- Be careful not to overwrite user changes outside the files touched for the rewrite.

