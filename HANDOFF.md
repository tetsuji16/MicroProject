# MicroProject Handoff

This repository is the shared workspace for two connected tracks:

1. A latest-snapshot import of ProjectLibre from SourceForge, kept under `upstream/projectlibre-snapshot`.
2. A Rust + Tauri shell that will first wrap the Java ProjectLibre code and later replace it slice by slice.

## Current Truth

- Repo purpose: `MicroProject` is the public home for the migration.
- Upstream provenance:
  - SourceForge tree: https://sourceforge.net/p/projectlibre/code/ci/master/tree/
  - Imported commit: `0530be227f4a10c5545cce8d3db20ac5a4d76a66`
- Licensing:
  - Upstream code remains CPAL 1.0 with its own notices inside the imported tree.
  - New migration work is documented separately at the repo root.
- Strategy:
  - Rust launches and orchestrates a Java adapter first.
  - The Java side is the behavior reference for `.mpp` import, compatibility, and ProjectLibre UI details.
  - Rust replaces Java only after each behavior slice is proven stable.

## Important Paths

- `README.md`: primary project overview and entry point.
- `JAVA_BRIDGE.md`: migration strategy and protocol shape.
- `NOTICE`: provenance and license summary.
- `ROADMAP.md`: phased implementation plan.
- `CONTRIBUTING.md`: how to contribute safely.
- `upstream/projectlibre-snapshot/`: imported ProjectLibre snapshot, treat as read-only unless refreshing upstream.
- `projectlibre-tauri/`: Rust + Tauri shell and UI scaffold.

## What Is Already Done

- Added root documentation and licensing files.
- Imported the ProjectLibre latest snapshot into `upstream/projectlibre-snapshot`.
- Built exploratory Rust/Tauri scaffolding that can be reused or pared back.
- Added ProjectLibre manual and sample `.mpp` files as compatibility references.
- `projectlibre-tauri` is still the place where the shell work should happen, but the first objective is bridge integration rather than a clean-room rewrite.

## Next Best Steps

1. Define the Java bridge process and IPC contract.
2. Add a Rust launcher that starts the Java adapter and proxies commands.
3. Make the bridge open the sample `.mpp` files and return a workspace snapshot.
4. Replace Java pieces with Rust only after the behavior matches the reference.

## Notes For Future AI

- Do not collapse the upstream snapshot into the rewrite code. Keep provenance visible.
- Keep `main` focused on the migration, not on full upstream history mirroring.
- Prefer incremental, testable vertical slices over a big-bang port.
- If the current Rust UI experiments conflict with the bridge-first strategy, treat them as exploratory and do not let them drive the architecture.
- Be careful not to overwrite user changes outside the files touched for the rewrite.
