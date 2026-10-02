# Handoff — continue Almadocx from here

**Last updated:** 2026-10-02  
**Repo:** https://github.com/almadocx/almadocx  
**Default branch:** `main` (PR #1 merged: `b2117b1`)  
**Primary fixture:** `fixtures/seamlesshr-golive-plan.docx` (~34 pages, many tables)

Open this repo in a new workspace/chat and say: *Resume from HANDOFF.md — continue M3 table parity.*

---

## Product goal

Canvas Word-like DOCX/ODT editor. Full editing, layout fidelity, UX parity with Microsoft Word.

This is a **pnpm monorepo**. `@almadocx/core` and `@almadocx/canvas` are workspace packages (`packages/*`), not external npm deps.

## Done on main (via #1)

- Typing perf: incremental `patchLayoutParagraph`, cached measurer, debounced reflow
- Undo/redo restores caret; marks are undoable
- Word-like bold/italic: range-only apply; clear sticky on caret move; inherit left-of-caret
- Table foundations: cell hit-test, selection over cell fills, in-cell drag-select, Ctrl+A = cell
- Fixes: bullet indent bounce, triple Backspace
- Roadmap: `MILESTONES.md`

## Next: M4+ (M3 largely shipped on `feat/m3-table-word-parity` / PR #3)

M3 table Word-parity is implemented (nav, rect select, structure ops, layout, cell typing patch, ARIA grid). Remaining polish: focus announcements, table style gallery.

Then M4 (styles), M5 (sections/headers), M6 (collab) as listed in `MILESTONES.md`.

## Architecture

| Package | Path | Role |
|---|---|---|
| `@almadocx/core` | `packages/core` | Model, ops/history, layout, DOCX/ODT |
| `@almadocx/canvas` | `packages/canvas` | Editor, hit-test, paint, IME |
| `@almadocx/playground` | `apps/playground` | Demo |
| `@almadocx/harness` | `apps/harness` | Layout regression |

## Process

- Branch from `main` (`fix/…`, `feat/…`); **PR only** — never push straight to `main`
- Push feature branch → `gh pr create` → merge when asked

## Run

```bash
git clone git@github.com:almadocx/almadocx.git
cd almadocx
pnpm install
pnpm --filter @almadocx/core test
pnpm --filter @almadocx/playground dev
```

Manual smoke: GoLive Plan → long-doc typing; bullets; marks; table columns + cell select + Ctrl+A.
