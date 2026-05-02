# Contributing

## How We Work

- Keep changes small and easy to review.
- Do not rewrite or reformat the imported upstream snapshot unless the change is strictly for provenance or licensing.
- Prefer one vertical slice at a time: model, storage, command, UI, test.

## Code Style

- Use Rust idioms and explicit types where they improve clarity.
- Keep new code ASCII-only unless a file already uses non-ASCII text.
- Add tests around persistence and dependency behavior as the rewrite grows.

## Pull Requests

- Summarize the user-visible change first.
- Call out whether the change touches upstream snapshot files or only rewrite code.
- Mention any manual verification steps you ran.

