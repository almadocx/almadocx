import { describe, expect, it, vi } from 'vitest'
import type {
  DocRange,
  LayoutPage,
  LayoutParagraph,
  LayoutResult,
  LayoutTable,
  LayoutTableCell,
} from '@almadocx/core'
import {
  cellsInRect,
  getCaretScreenRect,
  isRectCellSelection,
  lineSelectionSpan,
  moveByLayoutLine,
  offsetToX,
  paintSelection,
  paragraphIntersectsSelection,
  xToOffset,
} from '../src/selection/selectionPaint.js'

function fakePara(
  blockIndex: number,
  lines: Array<{ start: number; end: number; runs?: LayoutParagraph['lines'][number]['runs'] }>,
  cell?: LayoutParagraph['cell'],
): LayoutParagraph {
  return {
    kind: 'paragraph',
    sectionIndex: 0,
    blockIndex,
    x: 0,
    y: 0,
    width: 100,
    height: lines.length * 20 || 20,
    cell,
    lines: lines.map((l, i) => ({
      y: i * 20,
      height: 20,
      baseline: 16,
      runs:
        l.runs ??
        [
          {
            text: 'x'.repeat(Math.max(0, l.end - l.start)),
            x: 0,
            width: Math.max(0, l.end - l.start) * 8,
            font: '12px serif',
            color: '#000',
            startOffset: l.start,
          },
        ],
      startOffset: l.start,
      endOffset: l.end,
    })),
  }
}

function fakeCell(rowIndex: number, cellIndex: number): LayoutTableCell {
  return {
    x: cellIndex * 100,
    y: rowIndex * 40,
    width: 100,
    height: 40,
    paragraphs: [],
    borderColor: '#000',
    rowIndex,
    cellIndex,
  }
}

function fakeTable(rows: number, cols: number, blockIndex = 3): LayoutTable {
  const cells: LayoutTableCell[] = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) cells.push(fakeCell(r, c))
  }
  return {
    kind: 'table',
    blockIndex,
    sectionIndex: 0,
    x: 0,
    y: 0,
    width: cols * 100,
    height: rows * 40,
    cells,
    startRow: 0,
    endRow: rows - 1,
  }
}

describe('selection precision', () => {
  it('does not include preceding paragraphs', () => {
    const start = { sectionIndex: 0, blockIndex: 2, offset: 0 }
    const end = { sectionIndex: 0, blockIndex: 2, offset: 5 }
    expect(paragraphIntersectsSelection(fakePara(0, [{ start: 0, end: 10 }]), start, end)).toBe(false)
    expect(paragraphIntersectsSelection(fakePara(1, [{ start: 0, end: 10 }]), start, end)).toBe(false)
    expect(paragraphIntersectsSelection(fakePara(2, [{ start: 0, end: 10 }]), start, end)).toBe(true)
  })

  it('clips to selected line offsets within a paragraph', () => {
    const para = fakePara(0, [
      { start: 0, end: 10 },
      { start: 10, end: 20 },
      { start: 20, end: 30 },
    ])
    const start = { sectionIndex: 0, blockIndex: 0, offset: 12 }
    const end = { sectionIndex: 0, blockIndex: 0, offset: 18 }
    expect(lineSelectionSpan(para, para.lines[0]!, start, end)).toBeUndefined()
    expect(lineSelectionSpan(para, para.lines[1]!, start, end)).toEqual({ selStart: 12, selEnd: 18 })
    expect(lineSelectionSpan(para, para.lines[2]!, start, end)).toBeUndefined()
  })

  it('selects within a table cell without including neighbor columns', () => {
    const cell0 = fakePara(5, [{ start: 0, end: 8 }], { row: 0, cell: 0, para: 0 })
    const cell1 = fakePara(5, [{ start: 0, end: 12 }], { row: 0, cell: 1, para: 0 })
    const start = { sectionIndex: 0, blockIndex: 5, offset: 1, cell: { row: 0, cell: 0, para: 0 } }
    const end = { sectionIndex: 0, blockIndex: 5, offset: 6, cell: { row: 0, cell: 0, para: 0 } }
    expect(paragraphIntersectsSelection(cell0, start, end)).toBe(true)
    expect(paragraphIntersectsSelection(cell1, start, end)).toBe(false)
    expect(lineSelectionSpan(cell0, cell0.lines[0]!, start, end)).toEqual({
      selStart: 1,
      selEnd: 6,
    })
  })

  it('rejects paragraphs outside section / after end block / before start cell', () => {
    const start = { sectionIndex: 1, blockIndex: 1, offset: 0, cell: { row: 1, cell: 1, para: 0 } }
    const end = { sectionIndex: 1, blockIndex: 2, offset: 4, cell: { row: 1, cell: 2, para: 0 } }
    const earlySection = fakePara(1, [{ start: 0, end: 1 }])
    earlySection.sectionIndex = 0
    expect(paragraphIntersectsSelection(earlySection, start, end)).toBe(false)

    const lateSection = fakePara(1, [{ start: 0, end: 1 }])
    lateSection.sectionIndex = 2
    expect(paragraphIntersectsSelection(lateSection, start, end)).toBe(false)

    const afterEnd = fakePara(3, [{ start: 0, end: 1 }])
    afterEnd.sectionIndex = 1
    expect(paragraphIntersectsSelection(afterEnd, start, end)).toBe(false)

    const beforeCell = fakePara(1, [{ start: 0, end: 1 }], { row: 0, cell: 0, para: 0 })
    beforeCell.sectionIndex = 1
    expect(paragraphIntersectsSelection(beforeCell, start, end)).toBe(false)

    // Cell bounds are checked on the start block only
    const afterCellOnStart = fakePara(1, [{ start: 0, end: 1 }], { row: 1, cell: 3, para: 0 })
    afterCellOnStart.sectionIndex = 1
    expect(paragraphIntersectsSelection(afterCellOnStart, start, end)).toBe(false)

    // Body para on a cell start block is ordered before cell paths (cmpCell undefined < defined)
    const bodyAtStart = fakePara(1, [{ start: 0, end: 1 }])
    bodyAtStart.sectionIndex = 1
    expect(paragraphIntersectsSelection(bodyAtStart, start, end)).toBe(false)

    // Body selection (no cells) intersects body paragraphs
    const bodyOnly = fakePara(0, [{ start: 0, end: 3 }])
    expect(
      paragraphIntersectsSelection(bodyOnly, { sectionIndex: 0, blockIndex: 0, offset: 0 }, {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 2,
      }),
    ).toBe(true)

    // Cell para vs body-only endpoints: cmpCell(defined, undefined) > 0 → excluded on start block
    const cellVsBody = fakePara(0, [{ start: 0, end: 1 }], { row: 0, cell: 0, para: 0 })
    expect(
      paragraphIntersectsSelection(
        cellVsBody,
        { sectionIndex: 0, blockIndex: 0, offset: 0 },
        { sectionIndex: 0, blockIndex: 0, offset: 1 },
      ),
    ).toBe(false)
  })

  it('returns undefined when lineSelectionSpan collapses to empty', () => {
    const para = fakePara(0, [{ start: 0, end: 10 }])
    const start = { sectionIndex: 0, blockIndex: 0, offset: 5 }
    const end = { sectionIndex: 0, blockIndex: 0, offset: 5 }
    expect(lineSelectionSpan(para, para.lines[0]!, start, end)).toBeUndefined()
  })

  it('uses full line bounds for middle paragraphs in a multi-block span', () => {
    const mid = fakePara(1, [{ start: 0, end: 10 }])
    const start = { sectionIndex: 0, blockIndex: 0, offset: 2 }
    const end = { sectionIndex: 0, blockIndex: 2, offset: 4 }
    expect(lineSelectionSpan(mid, mid.lines[0]!, start, end)).toEqual({
      selStart: 0,
      selEnd: 10,
    })
  })

  it('treats body para vs body endpoints as equal cells via cmpCell(0)', () => {
    // end.cell set but start/para cells undefined → cmpCell(undefined, undefined) == 0
    const body = fakePara(0, [{ start: 0, end: 5 }])
    expect(
      paragraphIntersectsSelection(
        body,
        { sectionIndex: 0, blockIndex: 0, offset: 0 },
        { sectionIndex: 0, blockIndex: 0, offset: 3, cell: { row: 0, cell: 0, para: 0 } },
      ),
    ).toBe(true)
  })
})

describe('rectangular multi-cell selection', () => {
  it('detects a multi-cell table selection', () => {
    const base = { sectionIndex: 0, blockIndex: 3, offset: 0 }
    expect(
      isRectCellSelection(
        { ...base, cell: { row: 0, cell: 0, para: 0 } },
        { ...base, cell: { row: 0, cell: 0, para: 1 } },
      ),
    ).toBe(false)
    expect(
      isRectCellSelection(
        { ...base, cell: { row: 0, cell: 0, para: 0 } },
        { ...base, cell: { row: 0, cell: 2, para: 0 } },
      ),
    ).toBe(true)
    expect(
      isRectCellSelection(
        { ...base, cell: { row: 0, cell: 0, para: 0 } },
        { ...base, cell: { row: 2, cell: 0, para: 0 } },
      ),
    ).toBe(true)
    expect(isRectCellSelection(base, { ...base, offset: 4 })).toBe(false)
    expect(
      isRectCellSelection(
        { ...base, sectionIndex: 0, cell: { row: 0, cell: 0, para: 0 } },
        { ...base, sectionIndex: 1, cell: { row: 0, cell: 1, para: 0 } },
      ),
    ).toBe(false)
  })

  it('returns every layout cell inside the inclusive rectangle', () => {
    const table = fakeTable(3, 3)
    const inside = cellsInRect(table, 0, 1, 1, 2).map((c) => `${c.rowIndex},${c.cellIndex}`)
    expect(inside.sort()).toEqual(['0,1', '0,2', '1,1', '1,2'])
  })

  it('normalizes reversed rectangle bounds', () => {
    const table = fakeTable(3, 3)
    const forward = cellsInRect(table, 0, 0, 2, 2)
    const reversed = cellsInRect(table, 2, 2, 0, 0)
    expect(reversed.length).toBe(forward.length)
    expect(forward.length).toBe(9)
  })
})

describe('offsetToX / xToOffset', () => {
  it('handles empty runs, images, empty text, and past-end offsets', () => {
    const emptyLine = {
      y: 0,
      height: 20,
      baseline: 16,
      runs: [] as LayoutParagraph['lines'][number]['runs'],
      startOffset: 0,
      endOffset: 0,
    }
    expect(offsetToX(emptyLine, 0)).toBe(0)
    expect(xToOffset(emptyLine, 10)).toBe(0)

    const line = {
      y: 0,
      height: 20,
      baseline: 16,
      startOffset: 0,
      endOffset: 4,
      runs: [
        {
          text: '',
          image: { mediaId: 'm', width: 10, height: 10 },
          x: 0,
          width: 20,
          font: '12px',
          color: '#000',
          startOffset: 0,
        },
        {
          text: '',
          x: 20,
          width: 0,
          font: '12px',
          color: '#000',
          startOffset: 1,
        },
        {
          text: 'ab',
          x: 20,
          width: 20,
          font: '12px',
          color: '#000',
          startOffset: 2,
        },
      ],
    }
    expect(offsetToX(line, 0)).toBe(0)
    expect(offsetToX(line, 1)).toBe(20)
    expect(offsetToX(line, 3)).toBe(30)
    expect(offsetToX(line, 99)).toBe(40)
    expect(xToOffset(line, -5)).toBe(0)
    expect(xToOffset(line, 10)).toBe(1)
    expect(xToOffset(line, 20)).toBe(1)
    expect(xToOffset(line, 30)).toBe(3)
    expect(xToOffset(line, 100)).toBe(4)

    // Empty-text (non-image) run as the hit target for xToOffset charLen branch
    const emptyTextLine = {
      y: 0,
      height: 20,
      baseline: 16,
      startOffset: 0,
      endOffset: 1,
      runs: [
        {
          text: '',
          x: 0,
          width: 12,
          font: '12px',
          color: '#000',
          startOffset: 0,
        },
      ],
    }
    expect(xToOffset(emptyTextLine, 6)).toBe(1)
    expect(xToOffset(emptyTextLine, 0)).toBe(0)

    // Zero-width run → ratio short-circuit; x left of run.x clamps via Math.max(0, …)
    const zeroWidth = {
      y: 0,
      height: 20,
      baseline: 16,
      startOffset: 0,
      endOffset: 2,
      runs: [
        {
          text: 'ab',
          x: 10,
          width: 0,
          font: '12px',
          color: '#000',
          startOffset: 0,
        },
      ],
    }
    expect(xToOffset(zeroWidth, 10)).toBe(0)
    expect(xToOffset(zeroWidth, 5)).toBe(0)

    const wide = {
      y: 0,
      height: 20,
      baseline: 16,
      startOffset: 0,
      endOffset: 4,
      runs: [
        {
          text: 'abcd',
          x: 10,
          width: 40,
          font: '12px',
          color: '#000',
          startOffset: 0,
        },
      ],
    }
    // x slightly left of run still matches first run (x <= run.x+width) and clamps ratio
    expect(xToOffset(wide, 8)).toBe(0)
  })
})

describe('moveByLayoutLine', () => {
  it('moves within a paragraph and returns undefined at edges / missing para', () => {
    const para = fakePara(0, [
      { start: 0, end: 5 },
      { start: 5, end: 10 },
      { start: 10, end: 15 },
    ])
    const layout: LayoutResult = {
      pages: [
        {
          index: 0,
          width: 200,
          height: 400,
          blocks: [para],
          paragraphs: [para],
        },
      ],
    }
    const down = moveByLayoutLine(layout, { sectionIndex: 0, blockIndex: 0, offset: 2 }, 1)
    expect(down?.offset).toBeGreaterThanOrEqual(5)
    const up = moveByLayoutLine(layout, { sectionIndex: 0, blockIndex: 0, offset: 7 }, -1, 16)
    expect(up?.offset).toBeLessThan(5)
    expect(moveByLayoutLine(layout, { sectionIndex: 0, blockIndex: 0, offset: 0 }, -1)).toBeUndefined()
    expect(moveByLayoutLine(layout, { sectionIndex: 0, blockIndex: 0, offset: 14 }, 1)).toBeUndefined()
    expect(moveByLayoutLine(layout, { sectionIndex: 0, blockIndex: 9, offset: 0 }, 1)).toBeUndefined()

    const emptyLines = fakePara(1, [])
    const layoutEmpty: LayoutResult = {
      pages: [{ index: 0, width: 100, height: 100, blocks: [emptyLines], paragraphs: [emptyLines] }],
    }
    expect(
      moveByLayoutLine(layoutEmpty, { sectionIndex: 0, blockIndex: 1, offset: 0 }, 1),
    ).toBeUndefined()

    const cellPara = fakePara(2, [{ start: 0, end: 3 }, { start: 3, end: 6 }], {
      row: 0,
      cell: 0,
      para: 0,
    })
    const layoutCell: LayoutResult = {
      pages: [{ index: 0, width: 100, height: 100, blocks: [cellPara], paragraphs: [cellPara] }],
    }
    const moved = moveByLayoutLine(
      layoutCell,
      { sectionIndex: 0, blockIndex: 2, offset: 1, cell: { row: 0, cell: 0, para: 0 } },
      1,
    )
    expect(moved?.cell).toEqual({ row: 0, cell: 0, para: 0 })
    expect(
      moveByLayoutLine(
        layoutCell,
        { sectionIndex: 0, blockIndex: 2, offset: 1, cell: { row: 0, cell: 1, para: 0 } },
        1,
      ),
    ).toBeUndefined()

    // Offset past line end that still sits on a non-last line updates lineIndex via fallback
    const mid = moveByLayoutLine(layout, { sectionIndex: 0, blockIndex: 0, offset: 5 }, 1)
    expect(mid?.offset).toBeGreaterThanOrEqual(5)
  })
})

describe('paintSelection', () => {
  function mockCtx() {
    return {
      fillStyle: '',
      fillRect: vi.fn(),
    } as unknown as CanvasRenderingContext2D & { fillRect: ReturnType<typeof vi.fn> }
  }

  it('no-ops for collapsed ranges', () => {
    const ctx = mockCtx()
    const para = fakePara(0, [{ start: 0, end: 5 }])
    const page: LayoutPage = { index: 0, width: 100, height: 100, blocks: [para], paragraphs: [para] }
    const sel: DocRange = {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 2 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 2 },
    }
    paintSelection(ctx, page, sel)
    expect(ctx.fillRect).not.toHaveBeenCalled()
  })

  it('paints rectangular multi-cell fills and skips non-matching tables', () => {
    const ctx = mockCtx()
    const table = fakeTable(2, 2, 3)
    const other = fakeTable(1, 1, 9)
    const body = fakePara(3, [{ start: 0, end: 1 }])
    const page: LayoutPage = {
      index: 0,
      width: 400,
      height: 400,
      blocks: [body, other, table],
      paragraphs: [body],
    }
    paintSelection(ctx, page, {
      anchor: { sectionIndex: 0, blockIndex: 3, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 3, offset: 0, cell: { row: 1, cell: 1, para: 0 } },
    })
    expect(ctx.fillRect).toHaveBeenCalledTimes(4)

    const ctx2 = mockCtx()
    const noBlocks: LayoutPage = {
      index: 0,
      width: 100,
      height: 100,
      blocks: undefined as unknown as LayoutPage['blocks'],
      paragraphs: [],
    }
    paintSelection(ctx2, noBlocks, {
      anchor: { sectionIndex: 0, blockIndex: 3, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 3, offset: 0, cell: { row: 0, cell: 1, para: 0 } },
    })
    expect(ctx2.fillRect).not.toHaveBeenCalled()
  })

  it('paints text selection spans on intersecting paragraphs', () => {
    const ctx = mockCtx()
    const para = fakePara(0, [
      { start: 0, end: 10 },
      { start: 10, end: 20 },
    ])
    const skip = fakePara(5, [{ start: 0, end: 5 }])
    const page: LayoutPage = {
      index: 0,
      width: 200,
      height: 200,
      blocks: [para, skip],
      paragraphs: [para, skip],
    }
    paintSelection(ctx, page, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 2 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 15 },
    })
    expect(ctx.fillRect.mock.calls.length).toBeGreaterThanOrEqual(1)

    // Partial line selection skips lines before the start offset
    const ctx3 = mockCtx()
    paintSelection(ctx3, page, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 12 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 18 },
    })
    expect(ctx3.fillRect).toHaveBeenCalled()
  })
})

describe('getCaretScreenRect', () => {
  it('maps caret to screen coordinates for body and cell paragraphs', () => {
    const para = fakePara(0, [
      { start: 0, end: 5 },
      { start: 5, end: 10 },
    ])
    const cellPara = fakePara(1, [{ start: 0, end: 4 }], { row: 0, cell: 0, para: 0 })
    const layout: LayoutResult = {
      pages: [
        {
          index: 0,
          width: 200,
          height: 400,
          blocks: [para, cellPara],
          paragraphs: [para, cellPara],
        },
      ],
    }
    const body = getCaretScreenRect(
      layout,
      {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 7 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 7 },
      },
      1,
      24,
      1,
      800,
    )
    expect(body?.pageIndex).toBe(0)
    expect(body!.bottom).toBeGreaterThan(body!.top)

    const inCell = getCaretScreenRect(
      layout,
      {
        anchor: { sectionIndex: 0, blockIndex: 1, offset: 2, cell: { row: 0, cell: 0, para: 0 } },
        focus: { sectionIndex: 0, blockIndex: 1, offset: 2, cell: { row: 0, cell: 0, para: 0 } },
      },
      2,
    )
    expect(inCell?.left).toBeGreaterThanOrEqual(0)

    expect(
      getCaretScreenRect(
        layout,
        {
          anchor: { sectionIndex: 0, blockIndex: 1, offset: 2, cell: { row: 0, cell: 1, para: 0 } },
          focus: { sectionIndex: 0, blockIndex: 1, offset: 2, cell: { row: 0, cell: 1, para: 0 } },
        },
        1,
      ),
    ).toBeUndefined()

    expect(
      getCaretScreenRect(
        layout,
        {
          anchor: { sectionIndex: 0, blockIndex: 9, offset: 0 },
          focus: { sectionIndex: 0, blockIndex: 9, offset: 0 },
        },
        1,
      ),
    ).toBeUndefined()

    // Matching paragraph but offset outside every line → fall through
    expect(
      getCaretScreenRect(
        layout,
        {
          anchor: { sectionIndex: 0, blockIndex: 0, offset: 99 },
          focus: { sectionIndex: 0, blockIndex: 0, offset: 99 },
        },
        1,
      ),
    ).toBeUndefined()

    // Caret exactly at end of last line
    const atEnd = getCaretScreenRect(
      layout,
      {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 10 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 10 },
      },
      1,
    )
    expect(atEnd?.pageIndex).toBe(0)

    // Body caret with cell-less para when focus has no cell (skip cell check)
    const bodyOnly = getCaretScreenRect(
      layout,
      {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 3 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 3 },
      },
      1,
    )
    expect(bodyOnly).toBeDefined()
  })
})
