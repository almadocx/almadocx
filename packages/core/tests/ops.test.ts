import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  createHistory,
  dispatch,
  paragraphPlainText,
  redo,
  undo,
} from '../src/index.js'

describe('ops + history', () => {
  it('inserts text and undoes', () => {
    let doc = createEmptyDocument()
    let history = createHistory()
    ;({ doc, history } = dispatch(doc, history, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'Hello',
    }))
    expect(paragraphPlainText(doc.sections[0]!.blocks[0]!)).toBe('Hello')
    ;({ doc, history } = undo(doc, history))
    expect(paragraphPlainText(doc.sections[0]!.blocks[0]!)).toBe('')
    ;({ doc, history } = redo(doc, history))
    expect(paragraphPlainText(doc.sections[0]!.blocks[0]!)).toBe('Hello')
  })

  it('undo/redo restores caret selection', () => {
    let doc = createEmptyDocument()
    let history = createHistory()
    const before = {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 0 },
    }
    const after = {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 5 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 5 },
    }
    ;({ doc, history } = dispatch(
      doc,
      history,
      {
        type: 'insertText',
        position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        text: 'Hello',
      },
      { before, after },
    ))
    const undone = undo(doc, history)
    expect(undone.selection).toEqual(before)
    doc = undone.doc
    history = undone.history
    const redone = redo(doc, history)
    expect(redone.selection).toEqual(after)
  })

  it('undo restores setMark styling', () => {
    let doc = createEmptyDocument()
    let history = createHistory()
    ;({ doc, history } = dispatch(doc, history, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'bold',
    }))
    ;({ doc, history } = dispatch(doc, history, {
      type: 'setMark',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 4 },
      },
      mark: { bold: true },
    }))
    expect(doc.sections[0]!.blocks[0]!.runs.some((r) => r.props.bold === true)).toBe(true)
    ;({ doc, history } = undo(doc, history))
    expect(doc.sections[0]!.blocks[0]!.runs.every((r) => !r.props.bold)).toBe(true)
  })

  it('deleteRange inverts via insert', () => {
    let doc = createEmptyDocument()
    const r1 = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'abcdef',
    })
    doc = r1.doc
    const r2 = applyOp(doc, {
      type: 'deleteRange',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 2 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 4 },
      },
    })
    expect(paragraphPlainText(r2.doc.sections[0]!.blocks[0]!)).toBe('abef')
    const back = applyOp(r2.doc, r2.applied.inverse)
    expect(paragraphPlainText(back.doc.sections[0]!.blocks[0]!)).toBe('abcdef')
  })

  it('split and merge paragraph round-trip', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'HelloWorld',
    }).doc
    const split = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 5 },
    })
    expect(split.doc.sections[0]!.blocks).toHaveLength(2)
    expect(paragraphPlainText(split.doc.sections[0]!.blocks[0]!)).toBe('Hello')
    expect(paragraphPlainText(split.doc.sections[0]!.blocks[1]!)).toBe('World')
    const merged = applyOp(split.doc, split.applied.inverse)
    expect(merged.doc.sections[0]!.blocks).toHaveLength(1)
    expect(paragraphPlainText(merged.doc.sections[0]!.blocks[0]!)).toBe('HelloWorld')
  })

  it('setMark bold', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'bold',
    }).doc
    doc = applyOp(doc, {
      type: 'setMark',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 4 },
      },
      mark: { bold: true },
    }).doc
    const runs = doc.sections[0]!.blocks[0]!.runs
    expect(runs.some((r) => r.props.bold === true)).toBe(true)
  })
})
