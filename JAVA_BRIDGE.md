# Java Bridge Strategy

MicroProject now follows a wrapper-first migration path:

1. Keep the original ProjectLibre Java code as the functional baseline.
2. Build a Rust + Tauri shell that launches and talks to a Java adapter process.
3. Gradually move stable subsystems from Java to Rust once behavior matches the reference.

## Why This Approach

- ProjectLibre already contains a large amount of domain logic, UI behavior, and MS Project interoperability.
- A direct full rewrite is risky and makes it harder to preserve compatibility.
- A bridge-first approach lets us compare behavior against the original application while we port features in small slices.

## Initial Responsibilities

- Rust/Tauri:
  - App window, process lifecycle, file picker integration, and host-side settings.
  - IPC routing between the UI and the Java backend.
  - Long-term replacement for Java UI components.
- Java:
  - ProjectLibre business logic, import/export, scheduling, and any UI behavior we have not yet ported.
  - Compatibility reference for `.mpp` and legacy ProjectLibre behaviors.

## Suggested Bridge Shape

- Start with a Java adapter process launched by Rust.
- Exchange commands and snapshots over JSON lines on stdin/stdout or a local socket.
- Keep the protocol simple. The first implemented commands are:
  - `ping`
  - `snapshot`
  - `open`
  - `import_mpp`
  - `export_mpp`
  - `quit`
- Rust currently compiles the bridge for Java 8 bytecode, launches it on demand, and seeds the process with sample `.mpp` files from the imported ProjectLibre snapshot.
- Rust now also owns the default bridge state in-process. Java is kept as an opt-in fallback behind `MICROPROJECT_USE_JAVA_BRIDGE=1`.
- The next protocol layer should map those bridge calls onto real ProjectLibre behaviors instead of echo responses.

## Porting Rule

- When a behavior is stable and tested through the bridge, reimplement that slice in Rust.
- Do not remove the Java path until the Rust version is verified against the sample `.mpp` files and the manual.
- Prefer vertical slices over subsystem rewrites.
