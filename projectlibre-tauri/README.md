ProjectLibre Rust + TAURI MVP (Plan)*
====================================

- This directory contains the initial skeleton for the Rust + TAURI port of ProjectLibre.
- MVP scope includes: project/task CRUD, dependency handling, local persistence, and simple Gantt data feeding to a frontend UI.
- The backend is Rust; the UI is a TAURI shell with a Web frontend.
- This file is a starting point and will be evolved as the project progresses.

Build notes
- Ensure Rust toolchain is installed.
- Run frontend with a bundler (e.g., npm or pnpm) from frontend/ and integrate with TAURI via src-tauri.

Conventions
- IPC API sketch in src-tauri/ will evolve as the UI requirements are clarified.
- Data models in Rust (serde) should be designed for easy serialization to JSON/SQLite.
