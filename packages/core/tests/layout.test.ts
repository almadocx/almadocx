import { describe, expect, it } from 'vitest'
import { applyOp, createEmptyDocument, layoutDocument, patchLayoutParagraph, type Table } from '../src/index.js'

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
