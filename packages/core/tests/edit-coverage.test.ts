import { describe, expect, it } from 'vitest'
import {
  applyOp,
  charDirection,
  createEmptyDocument,
  createHistory,
  dispatch,
  dispatchBatch,
  extractPlainRange,
  findAll,
  findNext,
  isRtlChar,
  moveByWord,
  moveCaretByArrow,
  moveEnd,
  moveHome,
  moveLeft,
  moveRight,
  moveTableCellTab,
  moveTableCellVertical,
  moveVertical,
  paragraphBaseDirection,
  replaceAll,
  replaceMatch,
  canRedo,
  canUndo,
  undo,
  redo,
  type Document,
  type Table,
} from '../src/index.js'
import { paragraphAt } from '../src/edit/findReplace.js'

function cell(text: string, paras?: string[]) {
  const texts = paras ?? [text]
  return {
    id: `c_${texts.join('_')}`,
    props: {},
    blocks: texts.map((t, i) => ({
      id: `p_${texts.join('_')}_${i}`,
      type: 'paragraph' as const,
      props: {},
      runs: [{ id: `r_${i}`, props: {}, content: { type: 'text' as const, text: t } }],
    })),
  }
}

function tableDoc(rows: string[][]): Document {
  const doc = createEmptyDocument()
  const table: Table = {
    id: 't1',
    type: 'table',
    props: {},
    gridCols: rows[0]!.map(() => 2000),
    rows: rows.map((r, ri) => ({
      id: `row${ri}`,
      props: {},
      cells: r.map((t) => cell(t)),
    })),
  }
  doc.sections[0]!.blocks = [
    {
      id: 'before',
      type: 'paragraph',
      props: {},
      runs: [{ id: 'rb', props: {}, content: { type: 'text', text: 'Before' } }],
    },
    table,
    {
      id: 'after',
      type: 'paragraph',
      props: {},
      runs: [{ id: 'ra', props: {}, content: { type: 'text', text: 'After' } }],
    },
  ]
  return doc
}

describe('bidi', () => {
  it('classifies directions and moves caret/words', () => {
    expect(charDirection('')).toBe('neutral')
    expect(charDirection(' ')).toBe('neutral')
    expect(charDirection('5')).toBe('neutral')
    expect(charDirection(',')).toBe('neutral')
    expect(charDirection('a')).toBe('ltr')
    expect(charDirection('א')).toBe('rtl')
    expect(charDirection('ا')).toBe('rtl')
    expect(charDirection('★')).toBe('neutral')
    expect(isRtlChar('ب')).toBe(true)
    expect(isRtlChar('x')).toBe(false)
    expect(paragraphBaseDirection('')).toBe('ltr')
    expect(paragraphBaseDirection('  42')).toBe('ltr')
    expect(paragraphBaseDirection('שלום')).toBe('rtl')
    expect(paragraphBaseDirection('hello')).toBe('ltr')
    expect(moveCaretByArrow('abc', 0, 'ArrowLeft')).toBe(0)
    expect(moveCaretByArrow('abc', 3, 'ArrowRight')).toBe(3)
    expect(moveCaretByArrow('abc', 2, 'ArrowLeft')).toBe(1)
    expect(moveCaretByArrow('abc', 1, 'ArrowRight')).toBe(2)
    expect(moveCaretByArrow('שלום', 2, 'ArrowLeft', 'rtl')).toBe(1)
    expect(moveByWord('  hello  world', 2, -1)).toBe(0)
    expect(moveByWord('hello', 5, 1)).toBe(5)
  })
})

describe('navigation', () => {
  it('moves within and across paragraphs and tables', () => {
    const doc = tableDoc([
      ['A', 'B'],
      ['C', 'D'],
    ])
    const inA = { sectionIndex: 0, blockIndex: 1, offset: 1, cell: { row: 0, cell: 0, para: 0 } }
    expect(moveLeft(doc, inA).offset).toBe(0)
    expect(moveRight(doc, { ...inA, offset: 0 }).offset).toBe(1)
    expect(moveRight(doc, { ...inA, offset: 1 }).cell).toEqual({ row: 0, cell: 1, para: 0 })
    expect(moveLeft(doc, { sectionIndex: 0, blockIndex: 1, offset: 0, cell: { row: 0, cell: 0, para: 0 } }).blockIndex).toBe(0)

    const lastCell = { sectionIndex: 0, blockIndex: 1, offset: 1, cell: { row: 1, cell: 1, para: 0 } }
    expect(moveRight(doc, lastCell).blockIndex).toBe(2)
    expect(moveLeft(doc, { sectionIndex: 0, blockIndex: 2, offset: 0 }).cell?.row).toBe(1)

    expect(moveHome(doc, inA).offset).toBe(0)
    expect(moveHome(doc, inA, true)).toEqual({ sectionIndex: 0, blockIndex: 0, offset: 0 })
    expect(moveEnd(doc, inA).offset).toBe(1)
    const atDocEnd = moveEnd(doc, inA, true)
    expect(atDocEnd.blockIndex).toBe(2)
    expect(atDocEnd.offset).toBe(5)

    const tabNext = moveTableCellTab(doc, inA, 1)
    expect(tabNext.position.cell).toEqual({ row: 0, cell: 1, para: 0 })
    const tabLast = moveTableCellTab(doc, lastCell, 1)
    expect(tabLast.needsNewRow).toBe(true)
    const tabBack = moveTableCellTab(doc, { ...inA, cell: { row: 0, cell: 1, para: 0 } }, -1)
    expect(tabBack.position.cell?.cell).toBe(0)
    const tabUpRow = moveTableCellTab(doc, { ...inA, cell: { row: 1, cell: 0, para: 0 } }, -1)
    expect(tabUpRow.position.cell?.row).toBe(0)
    expect(moveTableCellTab(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, 1).position.blockIndex).toBe(0)

    expect(moveTableCellVertical(doc, inA, 1).cell?.row).toBe(1)
    expect(moveTableCellVertical(doc, { ...inA, cell: { row: 1, cell: 0, para: 0 } }, -1).cell?.row).toBe(0)
    expect(moveVertical(doc, { sectionIndex: 0, blockIndex: 0, offset: 2 }, 1).blockIndex).toBe(1)
    expect(moveVertical(doc, inA, 1).cell?.row).toBe(1)
    expect(moveLeft(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).offset).toBe(0)
    // Ctrl+Right at paragraph end crosses into the next block (first table cell)
    expect(moveRight(doc, { sectionIndex: 0, blockIndex: 0, offset: 6 }, true).blockIndex).toBe(1)
    expect(moveRight(doc, { sectionIndex: 0, blockIndex: 0, offset: 6 }, true).cell?.row).toBe(0)
  })

  it('handles multi-para cells, vMerge skip, and table-first docs', () => {
    const doc = createEmptyDocument()
    const table: Table = {
      id: 'tm',
      type: 'table',
      props: {},
      gridCols: [2000, 2000],
      rows: [
        {
          id: 'r0',
          props: {},
          cells: [
            {
              ...cell('top'),
              props: { vMerge: 'restart' },
              blocks: [
                cell('p0').blocks[0]!,
                cell('p1').blocks[0]!,
              ],
            },
            cell('right'),
          ],
        },
        {
          id: 'r1',
          props: {},
          cells: [
            { ...cell('cont'), props: { vMerge: 'continue' } },
            cell('br'),
          ],
        },
        {
          id: 'r2',
          props: {},
          cells: [cell('bottom'), cell('br2')],
        },
      ],
    }
    doc.sections[0]!.blocks = [table]

    expect(moveHome(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell).toBeTruthy()
    expect(moveEnd(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell?.row).toBe(2)

    const multi = {
      sectionIndex: 0,
      blockIndex: 0,
      offset: 0,
      cell: { row: 0, cell: 0, para: 0 },
    }
    expect(moveRight(doc, { ...multi, offset: 2 }).cell?.para).toBe(1)
    expect(moveLeft(doc, { ...multi, cell: { row: 0, cell: 0, para: 1 }, offset: 0 }).cell?.para).toBe(0)
    expect(moveTableCellVertical(doc, multi, 1).cell?.para).toBe(1)
    expect(moveTableCellVertical(doc, { ...multi, cell: { row: 0, cell: 0, para: 1 } }, -1).cell?.para).toBe(0)
    // Skip continue vMerge into bottom row
    const skipped = moveTableCellVertical(doc, multi, 1)
    // from para0 go to para1 first; from last para go to next row skipping continue
    const fromLastPara = moveTableCellVertical(
      doc,
      { ...multi, cell: { row: 0, cell: 0, para: 1 }, offset: 0 },
      1,
    )
    expect(fromLastPara.cell?.row).toBe(2)

    // Adjacent tables
    const doc2 = createEmptyDocument()
    doc2.sections[0]!.blocks = [
      {
        id: 't0',
        type: 'table',
        props: {},
        gridCols: [2000],
        rows: [{ id: 'r', props: {}, cells: [cell('X')] }],
      },
      {
        id: 't1',
        type: 'table',
        props: {},
        gridCols: [2000],
        rows: [{ id: 'r', props: {}, cells: [cell('Y')] }],
      },
    ]
    const fromX = {
      sectionIndex: 0,
      blockIndex: 0,
      offset: 1,
      cell: { row: 0, cell: 0, para: 0 },
    }
    expect(moveRight(doc2, fromX).blockIndex).toBe(1)
    expect(moveLeft(doc2, { sectionIndex: 0, blockIndex: 1, offset: 0, cell: { row: 0, cell: 0, para: 0 } }).blockIndex).toBe(0)
    void skipped
  })

  it('word-moves and vertical across non-table blocks', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'hello world',
    }).doc
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 11 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
      text: 'next',
    }).doc
    expect(moveRight(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).offset).toBe(6)
    expect(moveLeft(doc, { sectionIndex: 0, blockIndex: 0, offset: 11 }, true).offset).toBe(6)
    expect(moveVertical(doc, { sectionIndex: 0, blockIndex: 0, offset: 2 }, 1).blockIndex).toBe(1)
    expect(moveVertical(doc, { sectionIndex: 0, blockIndex: 1, offset: 2 }, -1).blockIndex).toBe(0)
    expect(moveEnd(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).blockIndex).toBe(1)
  })
})

describe('find/replace coverage', () => {
  it('covers whole-word, case, multi-block and table extract', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'Foo foo food',
    }).doc
    expect(findAll(doc, { query: '' })).toEqual([])
    // caseSensitive: 'foo' matches the word and the prefix of 'food'
    expect(findAll(doc, { query: 'foo', caseSensitive: true })).toHaveLength(2)
    expect(findAll(doc, { query: 'Foo', caseSensitive: true })).toHaveLength(1)
    expect(findAll(doc, { query: 'foo', wholeWord: true })).toHaveLength(2)
    expect(findNext(doc, { query: 'foo' }, { sectionIndex: 0, blockIndex: 0, offset: 4 })?.range.anchor.offset).toBe(4)
    expect(findNext(doc, { query: 'zzz' }, { sectionIndex: 0, blockIndex: 0, offset: 0 })).toBeUndefined()

    const m = findAll(doc, { query: 'Foo' })[0]!
    const replaced = replaceMatch(doc, m, 'Bar')
    expect(replaced.caret.offset).toBe(3)
    const cleared = replaceMatch(doc, m, '')
    expect(cleared.caret.offset).toBe(0)
    expect(replaceAll(doc, { query: 'foo' }, 'x').count).toBeGreaterThan(0)

    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 3 },
    }).doc
    expect(
      extractPlainRange(doc, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 1, offset: 2 },
      }),
    ).toContain('\n')

    const tdoc = tableDoc([
      ['AA', 'BB'],
      ['CC', 'DD'],
    ])
    expect(
      extractPlainRange(tdoc, {
        anchor: { sectionIndex: 0, blockIndex: 1, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        focus: { sectionIndex: 0, blockIndex: 1, offset: 2, cell: { row: 1, cell: 1, para: 0 } },
      }),
    ).toContain('\t')
    // reversed selection
    expect(
      extractPlainRange(tdoc, {
        focus: { sectionIndex: 0, blockIndex: 1, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        anchor: { sectionIndex: 0, blockIndex: 1, offset: 2, cell: { row: 1, cell: 1, para: 0 } },
      }),
    ).toContain('AA')

    expect(paragraphAt(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 })?.type).toBe('paragraph')
    expect(paragraphAt(tdoc, { sectionIndex: 0, blockIndex: 1, offset: 0 })).toBeUndefined()
  })
})

describe('history batch', () => {
  it('batches undo/redo and no-ops empty stacks', () => {
    let doc = createEmptyDocument()
    let history = createHistory()
    expect(canUndo(history)).toBe(false)
    expect(canRedo(history)).toBe(false)
    expect(undo(doc, history).history).toEqual(history)
    expect(redo(doc, history).history).toEqual(history)

    ;({ doc, history } = dispatchBatch(doc, history, []))
    ;({ doc, history } = dispatchBatch(doc, history, [
      { type: 'insertText', position: { sectionIndex: 0, blockIndex: 0, offset: 0 }, text: 'ab' },
      { type: 'insertText', position: { sectionIndex: 0, blockIndex: 0, offset: 2 }, text: 'cd' },
    ]))
    expect(canUndo(history)).toBe(true)
    ;({ doc, history } = undo(doc, history))
    expect(doc.sections[0]!.blocks[0]!.runs[0]).toBeTruthy()
    ;({ doc, history } = redo(doc, history))
    expect(canRedo(history)).toBe(false)

    // single-op batch collapses
    ;({ doc, history } = dispatchBatch(doc, history, [
      { type: 'insertText', position: { sectionIndex: 0, blockIndex: 0, offset: 4 }, text: '!' },
    ]))
    ;({ doc, history } = dispatch(
      doc,
      history,
      { type: 'insertText', position: { sectionIndex: 0, blockIndex: 0, offset: 5 }, text: '?' },
      {
        before: {
          anchor: { sectionIndex: 0, blockIndex: 0, offset: 5, cell: { row: 0, cell: 0, para: 0 } },
          focus: { sectionIndex: 0, blockIndex: 0, offset: 5, cell: { row: 0, cell: 0, para: 0 } },
        },
        after: {
          anchor: { sectionIndex: 0, blockIndex: 0, offset: 6, cell: { row: 0, cell: 0, para: 0 } },
          focus: { sectionIndex: 0, blockIndex: 0, offset: 6, cell: { row: 0, cell: 0, para: 0 } },
        },
      },
    ))
    ;({ doc, history } = undo(doc, history))
    expect(history.undoStack.length).toBeGreaterThan(0)
  })
})
