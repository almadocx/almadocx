# Handoff — continue Almadocx from here

**Last session:** 2026-10-02  
**Branch:** `fix/typing-perf-undo-selection`  
**Primary fixture:** `fixtures/seamlesshr-golive-plan.docx` (~34 pages, many tables)

Open this file in a new workspace/chat and tell the agent to resume from it.

---

## Product goal

Canvas Word-like DOCX/ODT editor (`~/Projects/almadocx`). Full editing, layout fidelity, UX parity with Microsoft Word.

## What landed this session

### Typing performance (priority #1 — done)
- Cached canvas measurer + measure-text cache
- `patchLayoutParagraph` fast path for insert/delete (no full reflow every keystroke)
- Debounced full layout + a11y `onChange`
- Caret/selection moves paint without layout
- List indent bug fixed: patch path was double-applying bullet indent every keystroke (“text bouncing right”)

### Undo / redo (done)
- `HistoryEntry` stores `before` / `after` selection
- Undo/redo restores caret
- `setMark` (bold/italic/etc.) is undoable

### Character styles — Word-like (done)
- Range bold/italic applies **only** to selection; does **not** sticky-persist
- Pending toggles clear when caret moves (click / arrows)
- Typing inherits style from character **to the left** of caret
- Collapsed Ctrl+B/I toggles sticky override until caret moves

### Tables — partial (M3 in progress)
- Hit-test: cell box first; horizontal containment when row paras share `y` (fixed “only last column”)
- Selection highlight paints **above** cell fills (was invisible under shading)
- Drag-select stays inside starting cell
- Ctrl+A in a cell selects **that cell’s** content only
- Full Word table parity is **not** done — see `MILESTONES.md` § M3

### Input bugs fixed
- Backspace deleted 3 chars: keydown was on textarea + host + document → single listener now

## Known remaining issues / next work

1. **M3 table parity** (current roadmap priority) — `MILESTONES.md`
   - Multi-cell / rectangular selection
   - Tab between cells; Up/Down by line metrics inside a cell
   - Merged cells (`gridSpan` / `vMerge`), borders, padding
   - Fast typing path inside cells (currently forces full layout)
2. Toolbar mark buttons should reflect active style at caret (UI not wired)
3. Pre-existing flaky/failing canvas test: `tests/sanitize.test.ts` “maps bold/italic from html”
4. Playground may be on **http://localhost:5174/** if 5173 is busy; Vite aliases load `packages/*/src` directly

## Architecture map

| Package | Role |
|---|---|
| `@almadocx/core` | Model, ops/history, layout (`layout.ts`, `patchLayoutParagraph`), DOCX/ODT |
| `@almadocx/canvas` | `editor.ts`, hit-test, selection paint, IME, canvas paint |
| `@almadocx/playground` | Demo app (`pnpm playground` / filter `@almadocx/playground`) |

Key files touched recently:
- `packages/canvas/src/editor.ts` — typing path, history selection, marks, cell select-all, single keydown
- `packages/canvas/src/selection/hitTest.ts` — table cell targeting
- `packages/canvas/src/selection/selectionPaint.ts` — selection geometry
- `packages/canvas/src/render/paint.ts` — paint order (table chrome → selection → text)
- `packages/core/src/ops/history.ts` — selection on undo/redo
- `packages/core/src/layout/layout.ts` — patch + list indent fix
- `MILESTONES.md` — roadmap (M3 tables)

## How to run

```bash
pnpm install
pnpm --filter @almadocx/core test
pnpm --filter @almadocx/canvas test   # sanitize bold/italic test may still fail (unrelated)
pnpm --filter @almadocx/core build && pnpm --filter @almadocx/canvas build
pnpm --filter @almadocx/playground dev
```

Manual checks:
1. Open GoLive Plan → type on long body text (should feel snappy)
2. Bullet line: type without text walking right; Backspace deletes one char
3. Select a word → italic → click into plain text → type (should be plain)
4. Click next to bold → type (should be bold)
5. Table ~page 22: click each column; drag-select; Ctrl+A selects cell only

## Agent / process notes

- Almasix repos: **PRs only** — never push to `main`; use `fix/` / `feat/` / `chore/` branches
- This repo previously had **no commits**; first commit may be on `fix/typing-perf-undo-selection`
- Do not recreate conversation history; use this file + `MILESTONES.md` + git log
