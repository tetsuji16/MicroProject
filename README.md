# MicroProject

MicroProject is the public home for a ProjectLibre latest-snapshot import and the ongoing Rust + Tauri rewrite.

## Current State

- Upstream reference snapshot: ProjectLibre `master` from SourceForge at commit `0530be227f4a10c5545cce8d3db20ac5a4d76a66`.
- Imported source lives in [`upstream/projectlibre-snapshot`](./upstream/projectlibre-snapshot).
- Rewrite scaffold lives in [`projectlibre-tauri`](./projectlibre-tauri).
- The repo is intentionally split so upstream provenance stays visible while the new app evolves independently.

## What This Repo Contains

- `upstream/projectlibre-snapshot/`: imported ProjectLibre snapshot for reference and provenance.
- `projectlibre-tauri/`: Rust + Tauri application scaffold that will become the new desktop app.
- `ROADMAP.md`: phase-by-phase rewrite plan.
- `NOTICE`: licensing and provenance notes for the combined repository.

## Getting Started

### Prerequisites

- Rust toolchain
- A desktop build environment for your platform
- Node.js only if you later add a bundler-based frontend

### Run the rewrite scaffold

```powershell
cd projectlibre-tauri
cargo run
```

## Development Approach

- Keep upstream snapshot changes isolated from rewrite work.
- Build the new app in small vertical slices: domain model, persistence, Tauri commands, then UI.
- Preserve file-level CPAL headers inside the imported snapshot.

## Roadmap

See [`ROADMAP.md`](./ROADMAP.md) for the implementation phases and next steps.

