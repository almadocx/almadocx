import { describe, expect, it } from 'vitest'
import {
  lineSelectionSpan,
  paragraphIntersectsSelection,
} from '../src/selection/selectionPaint.js'
import type { LayoutParagraph } from '@almadocx/core'

function fakePara(blockIndex: number, lines: Array<{ start: number; end: number }>): LayoutParagraph {
  return {
    kind: 'paragraph',
    sectionIndex: 0,
    blockIndex,
    x: 0,
    y: 0,
    width: 100,
    height: 40,
    lines: lines.map((l, i) => ({
      y: i * 20,
      height: 20,
      baseline: 16,
      runs: [{ text: 'x'.repeat(l.end - l.start), x: 0, width: (l.end - l.start) * 8, font: '12px serif', color: '#000', startOffset: l.start }],
      startOffset: l.start,
      endOffset: l.end,
    })),
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
    const cell0 = fakePara(5, [{ start: 0, end: 8 }])
    cell0.cell = { row: 0, cell: 0, para: 0 }
    const cell1 = fakePara(5, [{ start: 0, end: 12 }])
    cell1.cell = { row: 0, cell: 1, para: 0 }
    const start = { sectionIndex: 0, blockIndex: 5, offset: 1, cell: { row: 0, cell: 0, para: 0 } }
    const end = { sectionIndex: 0, blockIndex: 5, offset: 6, cell: { row: 0, cell: 0, para: 0 } }
    expect(paragraphIntersectsSelection(cell0, start, end)).toBe(true)
    expect(paragraphIntersectsSelection(cell1, start, end)).toBe(false)
    expect(lineSelectionSpan(cell0, cell0.lines[0]!, start, end)).toEqual({
      selStart: 1,
      selEnd: 6,
    })
  })
})
