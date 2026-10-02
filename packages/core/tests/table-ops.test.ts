import { beforeEach, describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  extractPlainRange,
  resetIdsForTests,
} from '../src/index.js'
import type { Document, Table } from '../src/index.js'

function cell(text: string, extraProps: Record<string, unknown> = {}) {
  return {
    id: `tc_${text}`,
    props: { ...extraProps },
    blocks: [
      {
        id: `p_${text}`,
        type: 'paragraph' as const,
        props: {},
        runs: [{ id: `r_${text}`, props: {}, content: { type: 'text' as const, text } }],
      },
    ],
  }
}

/** 2x2 table: rows [A,B] / [C,D], gridCols [3000, 3000]. */
function makeTableDoc(): Document {
  const doc = createEmptyDocument()
  doc.sections[0]!.blocks = [
    {
      id: 't1',
      type: 'table',
      props: {},
      gridCols: [3000, 3000],
      rows: [
        { id: 'row0', props: {}, cells: [cell('A'), cell('B')] },
        { id: 'row1', props: {}, cells: [cell('C'), cell('D')] },
      ],
    },
  ]
  return doc
}

function table(doc: Document): Table {
  const block = doc.sections[0]!.blocks[0]!
  if (block.type !== 'table') throw new Error('expected table')
  return block
}

/** Grid of cell text for easy comparison. */
function grid(doc: Document): string[][] {
  return table(doc).rows.map((r) =>
    r.cells.map((c) => {
      const p = c.blocks[0]
      return p && p.type === 'paragraph' ? p.runs.map((run) =>
        run.content.type === 'text' ? run.content.text : '').join('') : ''
    }),
  )
}

describe('table structural ops', () => {
  beforeEach(() => {
    resetIdsForTests()
  })

  describe('insertRow', () => {
    it('inserts an empty row after the given row', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'insertRow',
        sectionIndex: 0,
        blockIndex: 0,
        afterRow: 0,
      })
      expect(table(next).rows).toHaveLength(3)
      expect(grid(next)).toEqual([
        ['A', 'B'],
        ['', ''],
        ['C', 'D'],
      ])
    })

    it('inserts before the first row when afterRow is -1', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'insertRow',
        sectionIndex: 0,
        blockIndex: 0,
        afterRow: -1,
      })
      expect(grid(next)).toEqual([
        ['', ''],
        ['A', 'B'],
        ['C', 'D'],
      ])
    })

    it('preserves gridSpan structure of the template row', () => {
      const doc = makeTableDoc()
      // Make row0's first cell span 2 columns.
      table(doc).rows[0]!.cells = [cell('A', { gridSpan: 2 })]
      const { doc: next } = applyOp(doc, {
        type: 'insertRow',
        sectionIndex: 0,
        blockIndex: 0,
        afterRow: 0,
      })
      const inserted = table(next).rows[1]!
      expect(inserted.cells).toHaveLength(1)
      expect(inserted.cells[0]!.props.gridSpan).toBe(2)
    })

    it('round-trips via inverse', () => {
      const doc = makeTableDoc()
      const before = grid(doc)
      const { doc: next, applied } = applyOp(doc, {
        type: 'insertRow',
        sectionIndex: 0,
        blockIndex: 0,
        afterRow: 0,
      })
      expect(table(next).rows).toHaveLength(3)
      const { doc: restored } = applyOp(next, applied.inverse)
      expect(table(restored).rows).toHaveLength(2)
      expect(grid(restored)).toEqual(before)
    })
  })

  describe('deleteRow', () => {
    it('deletes the given row', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'deleteRow',
        sectionIndex: 0,
        blockIndex: 0,
        row: 0,
      })
      expect(grid(next)).toEqual([['C', 'D']])
    })

    it('refuses to delete the only remaining row', () => {
      const doc = makeTableDoc()
      table(doc).rows = [table(doc).rows[0]!]
      expect(() =>
        applyOp(doc, { type: 'deleteRow', sectionIndex: 0, blockIndex: 0, row: 0 }),
      ).toThrow(/only row/)
    })

    it('round-trips via inverse, restoring content verbatim', () => {
      const doc = makeTableDoc()
      const before = grid(doc)
      const { doc: next, applied } = applyOp(doc, {
        type: 'deleteRow',
        sectionIndex: 0,
        blockIndex: 0,
        row: 0,
      })
      expect(grid(next)).toEqual([['C', 'D']])
      const { doc: restored } = applyOp(next, applied.inverse)
      expect(grid(restored)).toEqual(before)
    })
  })

  describe('insertColumn', () => {
    it('inserts an empty cell after the given column in every row', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'insertColumn',
        sectionIndex: 0,
        blockIndex: 0,
        afterCol: 0,
      })
      expect(grid(next)).toEqual([
        ['A', '', 'B'],
        ['C', '', 'D'],
      ])
    })

    it('updates gridCols with the average width', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'insertColumn',
        sectionIndex: 0,
        blockIndex: 0,
        afterCol: 0,
      })
      expect(table(next).gridCols).toEqual([3000, 3000, 3000])
    })

    it('inserts before the first column when afterCol is -1', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'insertColumn',
        sectionIndex: 0,
        blockIndex: 0,
        afterCol: -1,
      })
      expect(grid(next)).toEqual([
        ['', 'A', 'B'],
        ['', 'C', 'D'],
      ])
    })

    it('round-trips via inverse', () => {
      const doc = makeTableDoc()
      const before = grid(doc)
      const beforeGrid = table(doc).gridCols
      const { doc: next, applied } = applyOp(doc, {
        type: 'insertColumn',
        sectionIndex: 0,
        blockIndex: 0,
        afterCol: 0,
      })
      expect(grid(next)[0]).toHaveLength(3)
      const { doc: restored } = applyOp(next, applied.inverse)
      expect(grid(restored)).toEqual(before)
      expect(table(restored).gridCols).toEqual(beforeGrid)
    })
  })

  describe('deleteColumn', () => {
    it('deletes the given column in every row', () => {
      const doc = makeTableDoc()
      const { doc: next } = applyOp(doc, {
        type: 'deleteColumn',
        sectionIndex: 0,
        blockIndex: 0,
        col: 0,
      })
      expect(grid(next)).toEqual([['B'], ['D']])
      expect(table(next).gridCols).toEqual([3000])
    })

    it('refuses to delete the only remaining column', () => {
      const doc = makeTableDoc()
      for (const r of table(doc).rows) r.cells = [r.cells[0]!]
      table(doc).gridCols = [3000]
      expect(() =>
        applyOp(doc, { type: 'deleteColumn', sectionIndex: 0, blockIndex: 0, col: 0 }),
      ).toThrow(/only column/)
    })

    it('round-trips via inverse, restoring content and gridCols verbatim', () => {
      const doc = makeTableDoc()
      const before = grid(doc)
      const beforeGrid = table(doc).gridCols
      const { doc: next, applied } = applyOp(doc, {
        type: 'deleteColumn',
        sectionIndex: 0,
        blockIndex: 0,
        col: 1,
      })
      expect(grid(next)).toEqual([['A'], ['C']])
      const { doc: restored } = applyOp(next, applied.inverse)
      expect(grid(restored)).toEqual(before)
      expect(table(restored).gridCols).toEqual(beforeGrid)
    })
  })
})

describe('extractPlainRange within a table', () => {
  it('slices offsets within a single cell paragraph', () => {
    const doc = makeTableDoc()
    table(doc).rows[0]!.cells[0]!.blocks[0] = {
      id: 'p',
      type: 'paragraph',
      props: {},
      runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'Hello' } }],
    }
    const text = extractPlainRange(doc, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 1, cell: { row: 0, cell: 0, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 4, cell: { row: 0, cell: 0, para: 0 } },
    })
    expect(text).toBe('ell')
  })

  it('joins multiple paragraphs in the same cell with \\n', () => {
    const doc = makeTableDoc()
    table(doc).rows[0]!.cells[0]!.blocks = [
      { id: 'p1', type: 'paragraph', props: {}, runs: [{ id: 'r1', props: {}, content: { type: 'text', text: 'one' } }] },
      { id: 'p2', type: 'paragraph', props: {}, runs: [{ id: 'r2', props: {}, content: { type: 'text', text: 'two' } }] },
    ]
    const text = extractPlainRange(doc, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 3, cell: { row: 0, cell: 0, para: 1 } },
    })
    expect(text).toBe('one\ntwo')
  })

  it('joins cells with \\t and rows with \\n (Word-like)', () => {
    const doc = makeTableDoc()
    const text = extractPlainRange(doc, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 1, cell: { row: 1, cell: 1, para: 0 } },
    })
    expect(text).toBe('A\tB\nC\tD')
  })

  it('extracts a partial multi-cell selection', () => {
    const doc = makeTableDoc()
    const text = extractPlainRange(doc, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 1, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 1, cell: { row: 1, cell: 0, para: 0 } },
    })
    expect(text).toBe('B\nC')
  })

  it('handles reversed selection direction', () => {
    const doc = makeTableDoc()
    const text = extractPlainRange(doc, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 1, cell: { row: 1, cell: 1, para: 0 } },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
    })
    expect(text).toBe('A\tB\nC\tD')
  })
})
