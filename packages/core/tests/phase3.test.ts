import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  findAll,
  replaceAll,
  moveLeft,
  moveRight,
  moveByWord,
  nextTabStopTwips,
  extractPlainRange,
} from '../src/index.js'

describe('phase 3 find/replace and navigation', () => {
  it('finds and replaces all', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'foo bar foo',
    }).doc
    const matches = findAll(doc, { query: 'foo' })
    expect(matches).toHaveLength(2)
    const { doc: next, count } = replaceAll(doc, { query: 'foo' }, 'baz')
    expect(count).toBe(2)
    expect(extractPlainRange(next, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 20 },
    })).toBe('baz bar baz')
  })

  it('moves by word', () => {
    expect(moveByWord('hello world', 0, 1)).toBe(6)
    expect(moveByWord('hello world', 11, -1)).toBe(6)
  })

  it('arrow navigation across paragraphs', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'ab',
    }).doc
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 2 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
      text: 'cd',
    }).doc
    const atEnd = { sectionIndex: 0, blockIndex: 0, offset: 2 }
    const next = moveRight(doc, atEnd)
    expect(next).toEqual({ sectionIndex: 0, blockIndex: 1, offset: 0 })
    const back = moveLeft(doc, next)
    expect(back).toEqual({ sectionIndex: 0, blockIndex: 0, offset: 2 })
  })

  it('computes default tab stops', () => {
    const stop = nextTabStopTwips(0, undefined, 12240)
    expect(stop.position).toBe(720)
    const second = nextTabStopTwips(720, undefined, 12240)
    expect(second.position).toBe(1440)
  })

  it('insertTab op grows offset by 1', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'A',
    }).doc
    const r = applyOp(doc, {
      type: 'insertTab',
      position: { sectionIndex: 0, blockIndex: 0, offset: 1 },
    })
    const plain = extractPlainRange(r.doc, {
      anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      focus: { sectionIndex: 0, blockIndex: 0, offset: 10 },
    })
    expect(plain).toBe('A\t')
  })
})
