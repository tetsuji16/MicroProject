# ProjectLibre Rewrite Scaffold

This directory contains the Rust + Tauri application scaffold for MicroProject.

## What It Does Today

- Stores workspace state locally as JSON.
- Exposes Tauri commands for project, task, dependency, and calendar CRUD.
- Keeps the frontend intentionally small while the domain model solidifies.

## Build

```powershell
cd projectlibre-tauri
cargo run
```

## Notes

- The frontend is still a placeholder and will be replaced by the real rewrite UI.
- The data model is JSON-first so we can iterate quickly before moving to richer persistence.
- The Tauri config lives in `src-tauri/tauri.conf.json`.
