# Almadocx

[![CI](https://img.shields.io/github/actions/workflow/status/almadocx/almadocx/ci.yml?branch=main&style=for-the-badge&logo=githubactions&logoColor=white&label=CI)](https://github.com/almadocx/almadocx/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/badge/coverage-100%25-brightgreen?style=for-the-badge&logo=vitest&logoColor=white)](./packages/core/vitest.config.ts)
[![Node](https://img.shields.io/badge/node-%3E%3D24-339933?style=for-the-badge&logo=nodedotjs&logoColor=white)](./package.json)
[![pnpm](https://img.shields.io/badge/pnpm-11-F69220?style=for-the-badge&logo=pnpm&logoColor=white)](./package.json)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178C6?style=for-the-badge&logo=typescript&logoColor=white)](./package.json)
[![DOCX](https://img.shields.io/badge/DOCX-round--trip-2B579A?style=for-the-badge&logo=microsoftword&logoColor=white)](./packages/core)
[![ODT](https://img.shields.io/badge/ODT-round--trip-18A303?style=for-the-badge&logo=libreoffice&logoColor=white)](./packages/core)

Canvas-based document editor with **DOCX** and **ODT** import/export, a format-neutral core, invertible editing operations, and a fidelity-first layout pipeline — aimed at Microsoft Word–class interaction.

## Why Almadocx

- **Canvas rendering** — pixel-stable pages, caret, and selection (not contenteditable hacks)
- **Invertible ops** — every edit has an inverse; undo/redo restores selection
- **Secure I/O** — zip bombs, path traversal, and XXE are treated as hostile by default ([SECURITY.md](./SECURITY.md))
- **Table parity** — cell hit-test, rectangular multi-cell selection, Tab/arrow nav, insert/delete row & column, ARIA grid
- **Snappy typing** — incremental layout patches for body and table cells on long documents

## Packages

| Package | Role |
|---|---|
| [`@almadocx/core`](./packages/core) | Model, styles, ops/history, layout, secure DOCX/ODT I/O |
| [`@almadocx/canvas`](./packages/canvas) | Canvas renderer, IME input, selection, ARIA mirror |
| [`@almadocx/playground`](./apps/playground) | Local demo app |
| [`@almadocx/harness`](./apps/harness) | Layout / regression harness |

## Quick start

```bash
pnpm install
pnpm playground          # Vite demo
pnpm test                # unit tests
pnpm test:coverage       # 100% coverage gate (core + canvas)
pnpm typecheck && pnpm lint
```

Requires **Node ≥ 24** and **pnpm 11**.

| Script | Purpose |
|---|---|
| `pnpm playground` | Run the playground editor |
| `pnpm test` | Run all package tests |
| `pnpm test:coverage` | Coverage with enforced 100% thresholds |
| `pnpm harness` | Layout stability harness |
| `pnpm typecheck` / `pnpm lint` | Static checks |

CI runs build → typecheck → lint → coverage → harness on every PR to `main` (see [`.github/workflows/ci.yml`](./.github/workflows/ci.yml)).

## Architecture (short)

```
DOCX/ODT ──► @almadocx/core (model + layout + ops)
                      │
                      ▼
              @almadocx/canvas (paint, IME, selection, a11y)
                      │
                      ▼
                 playground / host apps
```

## Status

| Milestone | State |
|---|---|
| M0–M2 | Done — model, canvas editor, long-doc typing perf |
| M2b–M3 | Done — table hit-test & Word-like table interaction |
| M4+ | Next — styling/marks parity, sections, collaboration |

Full roadmap: [MILESTONES.md](./MILESTONES.md). Continuing in a new chat? Start from [HANDOFF.md](./HANDOFF.md).

## Security

Document packages are **untrusted input**. See [SECURITY.md](./SECURITY.md) for zip, XML, and macro policy.
