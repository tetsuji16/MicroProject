# ProjectLibre Rewrite Scaffold

This directory contains the Rust + Tauri application scaffold for MicroProject.

## What It Does Today

- Stores workspace state locally as JSON through the Rust backend crate.
- Exposes Tauri commands for project, task, dependency, resource, assignment, and calendar CRUD.
- Supports XML import/export for workspace interchange.
- Keeps the frontend simple, but it now supports create, edit, delete, import, export, and recalculation flows.

## Build

```powershell
cd projectlibre-tauri
cargo run
```

## Notes

- The domain and persistence logic now lives in `backend/src/lib.rs`.
- The Tauri config lives in `src-tauri/tauri.conf.json`.
- The rewrite still has room to grow into richer ProjectLibre-compatible UI and file formats.
