import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  layoutDocument,
  patchLayoutCellParagraph,
  patchLayoutParagraph,
  type Table,
} from '../src/index.js'

describe('layout', () => {
  it('produces at least one page with lines', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'The quick brown fox jumps over the lazy dog. '.repeat(20),
    }).doc
    const layout = layoutDocument(doc)
    expect(layout.pages.length).toBeGreaterThanOrEqual(1)
    expect(layout.pages[0]!.paragraphs.length).toBeGreaterThanOrEqual(1)
    expect(layout.pages[0]!.paragraphs[0]!.lines.length).toBeGreaterThan(1)
  })

  it('lays out multi-column table cells with distinct x bands', () => {
    const doc = createEmptyDocument()
    const table: Table = {
      id: 't1',
      type: 'table',
      props: {},
      gridCols: [2000, 2000, 2000],
      rows: [
        {
          id: 'r1',
          props: {},
          cells: [0, 1, 2].map((i) => ({
            id: `c${i}`,
            props: {},
            blocks: [
              {
                id: `p${i}`,
                type: 'paragraph' as const,
                props: {},
                runs: [
                  {
                    id: `run${i}`,
                    props: {},
                    content: { type: 'text' as const, text: `Col ${i}` },
                  },
                ],
              },
            ],
          })),
        },
      ],
    }
    doc.sections[0]!.blocks = [table]
    const layout = layoutDocument(doc)
    const laid = layout.pages[0]!.blocks.find((b) => b.kind === 'table')
    expect(laid?.kind).toBe('table')
    if (laid?.kind !== 'table') return
    expect(laid.cells).toHaveLength(3)
    expect(laid.cells[0]!.x).toBeLessThan(laid.cells[1]!.x)
    expect(laid.cells[1]!.x).toBeLessThan(laid.cells[2]!.x)
    expect(laid.cells[0]!.width).toBeGreaterThan(10)
    expect(laid.cells[1]!.width).toBeGreaterThan(10)
    expect(laid.cells[2]!.width).toBeGreaterThan(10)
  })

  it('expands a vMerge restart cell to span continued rows', () => {
    const doc = createEmptyDocument()
    const makeCell = (id: string, text: string, vMerge?: 'restart' | 'continue') => ({
      id,
      props: vMerge ? { vMerge } : {},
      blocks: [
        {
          id: `${id}p`,
          type: 'paragraph' as const,
          props: {},
          runs: [{ id: `${id}r`, props: {}, content: { type: 'text' as const, text } }],
        },
      ],
    })
    const table: Table = {
      id: 'tm',
      type: 'table',
      props: {},
      gridCols: [2000, 2000],
      rows: [
        {
          id: 'row0',
          props: {},
          cells: [makeCell('a', 'merged', 'restart'), makeCell('b', 'top-right')],
        },
        {
          id: 'row1',
          props: {},
          cells: [makeCell('a2', '', 'continue'), makeCell('c', 'bottom-right')],
        },
      ],
    }
    doc.sections[0]!.blocks = [table]
    const layout = layoutDocument(doc)
    const laid = layout.pages[0]!.blocks.find((b) => b.kind === 'table')
    expect(laid?.kind).toBe('table')
    if (laid?.kind !== 'table') return
    const restart = laid.cells.find((c) => c.rowIndex === 0 && c.cellIndex === 0)!
    const topRight = laid.cells.find((c) => c.rowIndex === 0 && c.cellIndex === 1)!
    const bottomRight = laid.cells.find((c) => c.rowIndex === 1 && c.cellIndex === 1)!
    expect(restart).toBeTruthy()
    // continue cells are not laid out as content boxes
    expect(laid.cells.some((c) => c.rowIndex === 1 && c.cellIndex === 0)).toBe(false)
    // restart cell spans both rows
    expect(restart.height).toBeGreaterThan(topRight.height)
    expect(restart.height).toBeCloseTo(topRight.height + bottomRight.height, 1)
  })

  it('patchLayoutCellParagraph updates a cell paragraph in place', () => {
    const doc = createEmptyDocument()
    const table: Table = {
      id: 'tc',
      type: 'table',
      props: {},
      gridCols: [3000, 3000],
      rows: [
        {
          id: 'r0',
          props: {},
          cells: [0, 1].map((i) => ({
            id: `cell${i}`,
            props: {},
            blocks: [
              {
                id: `cp${i}`,
                type: 'paragraph' as const,
                props: {},
                runs: [{ id: `cr${i}`, props: {}, content: { type: 'text' as const, text: `Hi ${i}` } }],
              },
            ],
          })),
        },
      ],
    }
    doc.sections[0]!.blocks = [table]
    let layout = layoutDocument(doc)
    // Type a character into cell (0,0) paragraph 0 and patch in place.
    doc.sections[0]!.blocks = [doc.sections[0]!.blocks[0]!]
    const cellPara = (doc.sections[0]!.blocks[0] as Table).rows[0]!.cells[0]!.blocks[0]!
    cellPara.runs[0]!.content = { type: 'text', text: 'Hi 0 extra' }
    const patched = patchLayoutCellParagraph(doc, layout, 0, 0, { row: 0, cell: 0, para: 0 })
    layout = patched.layout
    const laid = layout.pages[0]!.blocks.find((b) => b.kind === 'table')
    expect(laid?.kind).toBe('table')
    if (laid?.kind !== 'table') return
    const cell = laid.cells.find((c) => c.rowIndex === 0 && c.cellIndex === 0)!
    const text = cell.paragraphs[0]!.lines.flatMap((l) => l.runs.map((r) => r.text)).join('')
    expect(text).toContain('extra')
  })

  it('patchLayoutParagraph keeps list indent stable across keystrokes', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'Item',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 0,
      numPr: { numId: '1', ilvl: 0 },
    }).doc
    let layout = layoutDocument(doc)
    const x0 = layout.pages[0]!.paragraphs[0]!.x
    expect(layout.pages[0]!.paragraphs[0]!.marker).toBeTruthy()

    for (let i = 0; i < 5; i++) {
      doc = applyOp(doc, {
        type: 'insertText',
        position: { sectionIndex: 0, blockIndex: 0, offset: 4 + i },
        text: 'x',
      }).doc
      const patched = patchLayoutParagraph(doc, layout, 0, 0)
      layout = patched.layout
      expect(layout.pages[0]!.paragraphs[0]!.x).toBeCloseTo(x0, 1)
    }
  })
})
