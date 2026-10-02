import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  layoutDocument,
  resolveListMarkers,
  saveDocument,
  loadDocument,
  resolveFontFamily,
} from '../src/index.js'

describe('phase 2 lists/tables/fonts', () => {
  it('resolves bullet and decimal markers', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'One',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 0,
      numPr: { numId: '1', ilvl: 0 },
    }).doc
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 3 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
      text: 'Two',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 1,
      numPr: { numId: '1', ilvl: 0 },
    }).doc
    const markers = resolveListMarkers(doc, 0)
    expect(markers.get(0)?.text).toContain('•')
    expect(markers.get(1)?.text).toContain('•')
    const layout = layoutDocument(doc)
    expect(layout.pages[0]!.paragraphs[0]!.marker?.text).toBeTruthy()
  })

  it('layouts a simple table', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [
      {
        id: 't1',
        type: 'table',
        props: {},
        gridCols: [3000, 3000],
        rows: [
          {
            id: 'r1',
            props: { header: true },
            cells: [
              {
                id: 'c1',
                props: {},
                blocks: [
                  {
                    id: 'p1',
                    type: 'paragraph',
                    props: {},
                    runs: [{ id: 'x', props: { bold: true }, content: { type: 'text', text: 'A' } }],
                  },
                ],
              },
              {
                id: 'c2',
                props: {},
                blocks: [
                  {
                    id: 'p2',
                    type: 'paragraph',
                    props: {},
                    runs: [{ id: 'y', props: {}, content: { type: 'text', text: 'B' } }],
                  },
                ],
              },
            ],
          },
        ],
      },
    ]
    const layout = layoutDocument(doc)
    const table = layout.pages[0]!.blocks.find((b) => b.kind === 'table')
    expect(table?.kind).toBe('table')
    if (table?.kind === 'table') expect(table.cells.length).toBe(2)
  })

  it('round-trips list numPr through DOCX', () => {
    let doc = createEmptyDocument('docx')
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'Listed',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 0,
      numPr: { numId: '2', ilvl: 0 },
    }).doc
    const bytes = saveDocument(doc, 'docx')
    const loaded = loadDocument(bytes, 'docx')
    const block = loaded.sections[0]!.blocks[0]!
    expect(block.type).toBe('paragraph')
    if (block.type === 'paragraph') {
      expect(block.props.numPr?.numId).toBeTruthy()
    }
  })

  it('substitutes Calibri to Carlito', () => {
    expect(resolveFontFamily('Calibri')).toBe('Carlito')
    expect(resolveFontFamily('Times New Roman')).toBe('Liberation Serif')
  })
})
