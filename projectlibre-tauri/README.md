# MicroProject Viewer

This directory contains the active Rust desktop viewer for Microsoft Project XML files.

## Preview

![MicroProject showcase](./frontend/demo-showcase.png)

## What It Does

- Opens Project XML files and renders a split task table and Gantt chart in `egui`.
- Parses the project header, calendars, tasks, predecessor links, and baseline fields into a forgiving view model.
- Uses native Rust crates for file dialogs, XML parsing, table layout, and rendering.
- Shows dependency lines and progress visuals directly on the chart so schedule structure is easy to read.

## Showcase Sample

- [`testdata/demo-showcase.xml`](./testdata/demo-showcase.xml) is a polished sample project with:
  - summary tasks
  - chained dependencies
  - milestones
  - progress percentages
  - baseline dates for visual comparison

Open it with:

```powershell
cargo run -- ".\testdata\demo-showcase.xml"
```

## Build And Run

```powershell
cd projectlibre-tauri
cargo run -- "C:\path\to\project.xml"
```

## Notes

- The viewer is Windows-first.
- The XML import is intentionally forgiving: unsupported fields are ignored rather than treated as fatal.
- The README at the repository root is the best place to explain the polished demo story for GitHub visitors.
- The legacy Tauri and bridge files are left in the tree only as historical artifacts.
