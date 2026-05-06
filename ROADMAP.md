# Roadmap

## Phase 0 - Strategy Reset

- Treat ProjectLibre Java as the authoritative behavior baseline.
- Keep the imported snapshot and the manual as the reference set for compatibility work.
- Use Rust/Tauri as the host shell, not as the first complete reimplementation.

## Phase 1 - Java Bridge

- Build a Java adapter process that can be launched by the Rust app.
- Define a small IPC contract for opening projects, importing `.mpp`, exporting `.mpp`, snapshotting workspace state, and triggering recalculation.
- Keep the bridge protocol JSON-based and easy to debug from the terminal.
- Wire the Rust shell to the adapter and verify it can open the sample `.mpp` files.

## Phase 2 - Compatibility Parity

- Compare the bridge output with the ProjectLibre manual and sample files.
- Align task, resource, calendar, baseline, and dependency behavior with the Java reference.
- Preserve ProjectLibre terminology in the UI and in the bridge protocol.
- Add regression tests around the sample `.mpp` files and XML interchange.

## Phase 3 - Rust Reimplementation

- Move stable domain objects from Java into Rust one slice at a time.
- Port persistence, scheduling, and import/export logic in the order that best preserves compatibility.
- Keep the Java bridge alive until the Rust replacement is verified against the reference behavior.
- Start by moving the bridge state, sample discovery, and command routing into Rust, then leave Java as an opt-in fallback.

## Phase 4 - Frontend Migration

- Replace any exploratory UI with screens that mirror the ProjectLibre workflow.
- Prefer editing forms and command flows that match the manual.
- Port view logic only after the corresponding bridge-backed behavior is stable.

## Phase 5 - Feature Growth

- Add calendar rules, task sorting, filters, grouping, and resource leveling.
- Expand coverage in small slices instead of a big-bang rewrite.
- Keep the sample `.mpp` files as recurring compatibility tests.
