# Almadocx

Canvas-based document editor with **DOCX** and **ODT** import/export, a format-neutral core, invertible editing operations, and a fidelity-first layout pipeline.

## Packages

| Package | Role |
|---|---|
| `@almadocx/core` | Model, styles, ops/history, layout, secure DOCX/ODT I/O |
| `@almadocx/canvas` | Canvas renderer, IME input, selection, ARIA mirror |
| `@almadocx/playground` | Product-quality demo app |
| `@almadocx/harness` | Layout/regression foundations |

## Develop

```bash
pnpm install
pnpm --filter @almadocx/core test
pnpm playground
```

## Security

See [SECURITY.md](./SECURITY.md). Packages are untrusted input (zip bombs, path traversal, XXE).

## Roadmap

See [MILESTONES.md](./MILESTONES.md) for the full plan. **Current priority: M3 — table interaction parity** (cell hit-test, selection, keyboard navigation, merged cells, Word-class UX).

**Continuing in a new chat/workspace?** Start from [HANDOFF.md](./HANDOFF.md).

## Phase 1 scope

Paragraphs, runs, styles cascade, pagination, caret/selection/IME, undo/redo, DOCX+ODT round-trip, ARIA mirror, polished playground.
