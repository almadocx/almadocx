# Almadocx milestones

Roadmap toward Word-class editing fidelity. Milestones are ordered; later work builds on earlier foundations.

## Completed (recent)

| ID | Milestone | Notes |
|---|---|---|
| M0 | Core model + DOCX/ODT round-trip | Paragraphs, runs, styles, sections |
| M1 | Canvas editor + IME + undo/redo | Caret, selection paint, history with selection restore |
| M2 | Long-document typing performance | Incremental paragraph layout, cached measurer, debounced reflow |
| M2b | Table hit-test foundation | Cell-aware targeting (fixes “always last column” bug) |

---

## M3 — Table interaction parity (current priority)

**Goal:** Tables feel like Microsoft Word: correct targeting, selection, navigation, and editing in every cell.

**Problem today:** Multi-column tables lay out every cell paragraph on the same baseline `y`. Naive hit-testing flattened all cell paragraphs and picked the last one in document order — so the caret and selection only worked reliably in the final column. Cell-first hit-testing is in place; the remaining work is broader table UX.

### M3.1 — Hit-test & pointer (in progress)

- [x] Resolve pointer to `(table, row, cell, para, offset)` via cell bounding box first
- [x] Prefer horizontal containment when row paragraphs share `y`
- [ ] Regression tests on GoLive fixture tables (all column counts, merged headers)
- [ ] Pointer coordinates under zoom / fit-width / horizontal scroll (sticky canvas)
- [ ] Double/triple-click word and paragraph selection inside cells

**Acceptance:** Click any column in a 5-column table on page 21+ of GoLive Plan; caret lands in that cell. Drag-select highlights text within the cell; Shift+arrow extends selection.

### M3.2 — Selection model

- [x] Single-cell drag selection highlight visible above cell fills
- [x] Ctrl+A in a cell selects that cell’s content only
- [ ] Multi-cell selection (drag across columns/rows — Word rectangular selection)
- [ ] Select row / select column (click gutter or table handle — future UI)
- [ ] Copy/cut/paste preserving table structure (HTML + plain + native fragment)
- [ ] Selection paint clipped to cell bounds (no bleed into neighbors)

**Acceptance:** Select “Module” in column 2 without selecting “Target Week” in column 4. Ctrl+A in a cell selects only that cell.

### M3.3 — Keyboard navigation

- [ ] Arrow keys move by character/line **within** cell using layout line metrics (not “next cell” on Up/Down)
- [ ] Tab / Shift+Tab move to next/previous cell (Word tab order)
- [ ] Home/End — cell-local vs table-wide (Ctrl+Home/End)
- [ ] Enter — new paragraph in cell vs exit table (configurable)
- [ ] Backspace at cell start — merge with previous cell content (not delete table)

**Acceptance:** Up/Down in a multi-line cell moves on the same column; Tab crosses columns left-to-right, row-wrap at row end.

### M3.4 — Layout fidelity

- [ ] `gridSpan` / `vMerge` — one layout cell per merged region; hit-test merged areas
- [ ] `tcW` / fixed column widths from DOCX when `tblGrid` is incomplete
- [ ] Row height: auto vs minimum from `trHeight`
- [ ] Cell padding (`tcMar`) and borders (`tcBorders`, `tblBorders`)
- [ ] Nested tables (stacking / recursion limits)

**Acceptance:** GoLive Plan tables render with correct column widths; merged header cells are one hit target.

### M3.5 — Editing operations

- [ ] Insert/delete characters in cell (fast path: patch cell paragraph layout)
- [ ] Split/merge paragraphs inside cell
- [ ] Insert/delete rows and columns (ops + UI)
- [ ] Table styles / autoformat (stretch)

**Acceptance:** Typing in any cell of a 34-page doc stays snappy; undo restores cell content and caret.

### M3.6 — Accessibility

- [ ] ARIA grid role with `row` / `gridcell` / active descendant
- [ ] Announce cell coordinates on focus change
- [ ] Screen reader selection and table navigation

---

## M4 — Styling & marks parity

- [ ] Undo/redo batches for compound formatting
- [ ] Character styles, paragraph styles in tables
- [ ] Lists inside table cells

## M5 — Page layout & sections

- [ ] Headers/footers editing
- [ ] Section breaks, columns, page fields
- [ ] Floating objects / images behind text

## M6 — Collaboration & polish

- [ ] Track changes (stretch)
- [ ] Comments (stretch)
- [ ] Print/PDF export matching canvas layout

---

## How to use this doc

1. Pick the next unchecked item in the **current priority** milestone (M3).
2. Add tests in `@almadocx/core` (layout/ops) or `@almadocx/canvas` (hit-test/selection) before closing each sub-milestone.
3. GoLive Plan (`fixtures/seamlesshr-golive-plan.docx`) is the primary regression doc for tables and long-document performance.

## Verification checklist (M3)

Run after table interaction changes:

```bash
pnpm --filter @almadocx/core test
pnpm --filter @almadocx/canvas test
pnpm playground
```

Manual:

1. Open GoLive Plan → navigate to a 4–5 column module table (~page 22).
2. Click each header column — caret must follow the click.
3. Drag to select text inside a single cell — blue highlight only in that cell.
4. Type — characters appear in the focused cell; undo restores caret position.
