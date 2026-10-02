import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  paragraphPlainText,
  type Document,
  type Table,
} from '../src/index.js'

function cell(text: string) {
  return {
    id: `c_${text}`,
    props: {},
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

function makeTableDoc(): Document {
  const doc = createEmptyDocument()
  doc.sections[0]!.blocks = [
    {
      id: 't1',
      type: 'table',
      props: {},
      gridCols: [2000, 2000],
      rows: [
        { id: 'r0', props: {}, cells: [cell('A'), cell('B')] },
        { id: 'r1', props: {}, cells: [cell('C'), cell('D')] },
      ],
    } satisfies Table,
  ]
  return doc
}

describe('ops apply coverage', () => {
  it('covers empty delete, multi-para delete, marks, lists, props, tabs', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'HelloWorld',
      props: { bold: true },
    }).doc

    // collapsed delete
    const emptyDel = applyOp(doc, {
      type: 'deleteRange',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 2 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 2 },
      },
    })
    expect(emptyDel.doc).toBe(doc)

    // setMark with italic/underline/strike invert
    doc = applyOp(doc, {
      type: 'setMark',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 5 },
      },
      mark: { italic: true, underline: true, strike: true },
    }).doc

    // split then multi-paragraph delete
    const split = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 5 },
    })
    doc = split.doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 5 },
      text: '!',
    }).doc
    // insert middle paragraph
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 5 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
      text: 'MID',
    }).doc
    const multi = applyOp(doc, {
      type: 'deleteRange',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 2 },
        focus: { sectionIndex: 0, blockIndex: 2, offset: 2 },
      },
    })
    expect(multi.doc.sections[0]!.blocks).toHaveLength(1)

    doc = createEmptyDocument()
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
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 0,
      numPr: null,
    }).doc
    expect(doc.sections[0]!.blocks[0]!.props.numPr).toBeUndefined()

    doc = applyOp(doc, {
      type: 'setParagraphProps',
      sectionIndex: 0,
      blockIndex: 0,
      props: { alignment: 'center', indentLeft: 200 },
    }).doc
    expect(doc.sections[0]!.blocks[0]!.props.alignment).toBe('center')

    // insertTab mid-run and at boundaries
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'AB',
    }).doc
    doc = applyOp(doc, {
      type: 'insertTab',
      position: { sectionIndex: 0, blockIndex: 0, offset: 1 },
    }).doc
    expect(paragraphPlainText(doc.sections[0]!.blocks[0]!)).toContain('\t')
    doc = applyOp(doc, {
      type: 'insertTab',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
    }).doc
  })

  it('covers table cell split/merge and insertTab after non-text', () => {
    let doc = makeTableDoc()
    const pos = {
      sectionIndex: 0,
      blockIndex: 0,
      offset: 1,
      cell: { row: 0, cell: 0, para: 0 },
    }
    const split = applyOp(doc, { type: 'splitParagraph', position: pos })
    doc = split.doc
    const cellBlocks = (doc.sections[0]!.blocks[0] as Table).rows[0]!.cells[0]!.blocks
    expect(cellBlocks).toHaveLength(2)
    doc = applyOp(doc, split.applied.inverse).doc
    expect((doc.sections[0]!.blocks[0] as Table).rows[0]!.cells[0]!.blocks).toHaveLength(1)

    // insertTab at offset 0 and after image-like atomic via tab at end of text then again
    doc = applyOp(doc, {
      type: 'insertTab',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
    }).doc
    // Put a tab run then insert another tab after it (offsetInRun !== 0 on non-text)
    const table = doc.sections[0]!.blocks[0] as Table
    const para = table.rows[0]!.cells[0]!.blocks[0]!
    para.runs = [
      { id: 't', props: {}, content: { type: 'tab' } },
      { id: 'x', props: {}, content: { type: 'text', text: 'Z' } },
    ]
    doc = applyOp(doc, {
      type: 'insertTab',
      position: { sectionIndex: 0, blockIndex: 0, offset: 1, cell: { row: 0, cell: 0, para: 0 } },
    }).doc
  })

  it('covers insertColumn without gridCols and delete with capture', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [
          { id: 'r0', props: {}, cells: [cell('A'), cell('B')] },
          { id: 'r1', props: {}, cells: [cell('C'), cell('D')] },
        ],
      },
    ]
    const inserted = applyOp(doc, {
      type: 'insertColumn',
      sectionIndex: 0,
      blockIndex: 0,
      afterCol: 0,
    })
    expect((inserted.doc.sections[0]!.blocks[0] as Table).rows[0]!.cells).toHaveLength(3)

    // with empty cells template path using afterCol -1
    const withGrid = makeTableDoc()
    const col = applyOp(withGrid, {
      type: 'insertColumn',
      sectionIndex: 0,
      blockIndex: 0,
      afterCol: -1,
      gridCol: 1500,
    })
    expect((col.doc.sections[0]!.blocks[0] as Table).gridCols?.[0]).toBe(1500)

    // deleteRange with pre-captured deletedText
    let d = createEmptyDocument()
    d = applyOp(d, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'xyz',
    }).doc
    d = applyOp(d, {
      type: 'deleteRange',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 3 },
      },
      deletedText: 'xyz',
      deletedProps: { bold: true },
    }).doc
    expect(paragraphPlainText(d.sections[0]!.blocks[0]!)).toBe('')
  })
})
