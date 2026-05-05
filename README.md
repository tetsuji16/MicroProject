# MicroProject

MicroProject is the public home for a ProjectLibre latest-snapshot import and a wrapper-first migration to Rust + Tauri.

## Current State

- Upstream reference snapshot: ProjectLibre `master` from SourceForge at commit `0530be227f4a10c5545cce8d3db20ac5a4d76a66`.
- Imported source lives in [`upstream/projectlibre-snapshot`](./upstream/projectlibre-snapshot).
- The migration target is a Rust + Tauri shell that first wraps the Java ProjectLibre code and then replaces it slice by slice.
- The wrapper strategy is documented in [`JAVA_BRIDGE.md`](./JAVA_BRIDGE.md).
- The current implementation is Rust-first: bridge state, sample discovery, and command routing live in Rust, while Java remains an opt-in fallback via `MICROPROJECT_USE_JAVA_BRIDGE=1`.
- The repo is intentionally split so upstream provenance stays visible while the new app evolves independently.

## What This Repo Contains

- `upstream/projectlibre-snapshot/`: imported ProjectLibre snapshot for reference and provenance.
- `projectlibre-tauri/`: Rust + Tauri shell and UI scaffold for the migration.
- `JAVA_BRIDGE.md`: the bridge-first migration strategy and protocol shape.
- `ROADMAP.md`: phase-by-phase migration plan.
- `NOTICE`: licensing and provenance notes for the combined repository.

## Getting Started

### Prerequisites

- Rust toolchain
- A Java runtime and build toolchain for the ProjectLibre bridge phase
- A desktop build environment for your platform
- Node.js only if you later add a bundler-based frontend

### Run the migration shell

```powershell
cd projectlibre-tauri
cargo run
```

## Development Approach

- Keep upstream snapshot changes isolated from migration work.
- Start with a Java adapter process that exposes ProjectLibre behavior to Rust through a small IPC contract.
- Use Rust/Tauri for the shell, windowing, and orchestration first.
- Replace Java subsystems only after the bridge behavior is verified against the manual and sample `.mpp` files.
- Preserve file-level CPAL headers inside the imported snapshot.

## Roadmap

See [`ROADMAP.md`](./ROADMAP.md) for the migration phases and next steps.
