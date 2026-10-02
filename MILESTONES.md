# Almadocx milestones

Roadmap toward Word-class editing fidelity. Milestones are ordered; later work builds on earlier foundations.

## Completed (recent)

| ID | Milestone | Notes |
|---|---|---|
| M0 | Core model + DOCX/ODT round-trip | Paragraphs, runs, styles, sections |
| M1 | Canvas editor + IME + undo/redo | Caret, selection paint, history with selection restore |
| M2 | Long-document typing performance | Incremental paragraph layout, cached measurer, debounced reflow |
| M2b | Table hit-test foundation | Cell-aware targeting (fixes “always last column” bug) |
| M3 | Table interaction parity | Cell nav, rect select, structure ops, layout fidelity, ARIA grid |

---

## M3 — Table interaction parity (largely complete)

**Goal:** Tables feel like Microsoft Word: correct targeting, selection, navigation, and editing in every cell.

### M3.1 — Hit-test & pointer

- [x] Resolve pointer to `(table, row, cell, para, offset)` via cell bounding box first
- [x] Prefer horizontal containment when row paragraphs share `y`
- [x] Regression tests on GoLive fixture tables
- [x] Double/triple-click word and paragraph selection inside cells

### M3.2 — Selection model

- [x] Single-cell drag selection highlight visible above cell fills
- [x] Ctrl+A in a cell selects that cell’s content only
- [x] Multi-cell rectangular selection (drag across columns/rows)
- [x] Ctrl+A escalation: cell → table → document
- [x] Copy/cut plain text from cells (`extractPlainRange`)
- [x] Selection paint for rect cells + text spans in one cell

### M3.3 — Keyboard navigation

- [x] Arrow Up/Down by visual layout line within paragraph; then same-column cell
- [x] Tab / Shift+Tab move to next/previous cell
- [x] Tab in last cell inserts a new row
- [x] Backspace at cell start moves to previous cell
- [x] Enter — new paragraph in cell

### M3.4 — Layout fidelity

- [x] `gridSpan` / `vMerge` — restart cells grow to cover continue rows
- [x] `tcMar` / `tcW` / borders / `tblPr` parse + serialize (best-effort)
- [x] Row height from content + `trHeight`
- [x] Cell padding from `tcMar`; `vAlign` offset

### M3.5 — Editing operations

- [x] Insert/delete characters in cell (fast path: `patchLayoutCellParagraph`)
- [x] Split/merge paragraphs inside cell
- [x] Insert/delete rows and columns (invertible ops)
- [ ] Table styles / autoformat (stretch — deferred)

### M3.6 — Accessibility

- [x] ARIA `role="grid"` with row / gridcell and indices
- [ ] Announce cell coordinates on focus change (polish)
- [ ] Screen reader selection and table navigation (polish)

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

## Verification checklist (M3)

```bash
pnpm --filter @almadocx/core test
pnpm --filter @almadocx/canvas test
pnpm playground
```

Manual:

1. Open GoLive Plan → ~page 22 multi-column table
2. Click each header column — caret follows
3. Drag-select text in one cell; drag across cells — rectangular highlight
4. Tab / Shift+Tab between cells; Tab at end adds a row
5. Type in a cell — snappy; Ctrl+C copies cell text; undo restores caret
