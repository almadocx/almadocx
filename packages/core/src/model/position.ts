import { AlmadocxError, assert } from '../util/assert.js'
import type { Document, Paragraph, Run, Table } from './types.js'

/** Path into a table cell paragraph when the block is a table. */
export interface CellPath {
  row: number
  cell: number
  para: number
}

/** Absolute character offset across the whole document (text + structural break points). */
export interface DocPosition {
  sectionIndex: number
  blockIndex: number
  /** UTF-16 offset within the paragraph's flattened text (tabs/breaks count as 1). */
  offset: number
  /** Present when the caret is inside a table cell. */
  cell?: CellPath
}

export interface DocRange {
  anchor: DocPosition
  focus: DocPosition
}

function compareCell(a: CellPath | undefined, b: CellPath | undefined): number {
  if (!a && !b) return 0
  if (!a) return -1
  if (!b) return 1
  if (a.row !== b.row) return a.row - b.row
  if (a.cell !== b.cell) return a.cell - b.cell
  return a.para - b.para
}

export function comparePositions(a: DocPosition, b: DocPosition): number {
  if (a.sectionIndex !== b.sectionIndex) return a.sectionIndex - b.sectionIndex
  if (a.blockIndex !== b.blockIndex) return a.blockIndex - b.blockIndex
  const cellCmp = compareCell(a.cell, b.cell)
  if (cellCmp !== 0) return cellCmp
  return a.offset - b.offset
}

export function positionsEqual(a: DocPosition, b: DocPosition): boolean {
  return comparePositions(a, b) === 0
}

export function normalizeRange(range: DocRange): { start: DocPosition; end: DocPosition } {
  return comparePositions(range.anchor, range.focus) <= 0
    ? { start: range.anchor, end: range.focus }
    : { start: range.focus, end: range.anchor }
}

export function isCollapsed(range: DocRange): boolean {
  return positionsEqual(range.anchor, range.focus)
}

/** Flatten paragraph inline content to a string used for offsets (tab/break → \t/\n). */
export function paragraphPlainText(paragraph: Paragraph): string {
  let out = ''
  for (const run of paragraph.runs) {
    out += inlineToChar(run)
  }
  return out
}

/** Object replacement character for inline images in plain-text offsets. */
export const IMAGE_PLACEHOLDER = '\uFFFC'

export function inlineToChar(run: Run): string {
  switch (run.content.type) {
    case 'text':
      return run.content.text
    case 'tab':
      return '\t'
    case 'break':
      return '\n'
    case 'image':
      return IMAGE_PLACEHOLDER
  }
}

export function paragraphLength(paragraph: Paragraph): number {
  return paragraphPlainText(paragraph).length
}

export function nearestParagraphIndex(doc: Document, sectionIndex: number, blockIndex: number): number {
  const section = doc.sections[sectionIndex]
  if (!section || section.blocks.length === 0) return 0
  if (section.blocks[blockIndex]?.type === 'paragraph') return blockIndex
  // Prefer entering a table at the caret rather than skipping past it
  if (section.blocks[blockIndex]?.type === 'table') return blockIndex
  for (let i = blockIndex; i < section.blocks.length; i++) {
    if (section.blocks[i]?.type === 'paragraph' || section.blocks[i]?.type === 'table') return i
  }
  for (let i = blockIndex; i >= 0; i--) {
    if (section.blocks[i]?.type === 'paragraph' || section.blocks[i]?.type === 'table') return i
  }
  return 0
}

function firstCellPath(table: Table): CellPath {
  for (let row = 0; row < table.rows.length; row++) {
    const cells = table.rows[row]?.cells ?? []
    for (let cell = 0; cell < cells.length; cell++) {
      const blocks = cells[cell]?.blocks ?? []
      if (blocks.length > 0) return { row, cell, para: 0 }
    }
  }
  return { row: 0, cell: 0, para: 0 }
}

export function getParagraph(doc: Document, pos: DocPosition): Paragraph {
  const section = doc.sections[pos.sectionIndex]
  assert(section, 'bad_position', `section ${String(pos.sectionIndex)} missing`)
  const bi = nearestParagraphIndex(doc, pos.sectionIndex, pos.blockIndex)
  const block = section.blocks[bi]
  assert(block, 'bad_position', `block ${String(bi)} missing`)
  if (block.type === 'table') {
    const cell = pos.cell ?? firstCellPath(block)
    const row = block.rows[cell.row]
    assert(row, 'bad_position', `table row ${String(cell.row)} missing`)
    const tc = row.cells[cell.cell]
    assert(tc, 'bad_position', `table cell ${String(cell.cell)} missing`)
    const para = tc.blocks[cell.para]
    assert(para?.type === 'paragraph', 'bad_position', 'expected cell paragraph')
    return para
  }
  assert(block.type === 'paragraph', 'bad_position', 'expected paragraph')
  return block
}

export function clampPosition(doc: Document, pos: DocPosition): DocPosition {
  const sectionIndex = Math.max(0, Math.min(pos.sectionIndex, doc.sections.length - 1))
  const section = doc.sections[sectionIndex]
  if (!section || section.blocks.length === 0) {
    return { sectionIndex: 0, blockIndex: 0, offset: 0 }
  }
  let blockIndex = Math.max(0, Math.min(pos.blockIndex, section.blocks.length - 1))
  blockIndex = nearestParagraphIndex(doc, sectionIndex, blockIndex)
  const block = section.blocks[blockIndex]
  if (!block) {
    return { sectionIndex, blockIndex, offset: 0 }
  }
  if (block.type === 'table') {
    const cell = pos.cell ?? firstCellPath(block)
    const row = block.rows[Math.max(0, Math.min(cell.row, block.rows.length - 1))]
    if (!row || row.cells.length === 0) {
      return { sectionIndex, blockIndex, offset: 0, cell: firstCellPath(block) }
    }
    const cellIndex = Math.max(0, Math.min(cell.cell, row.cells.length - 1))
    const tc = row.cells[cellIndex]!
    const paraIndex = Math.max(0, Math.min(cell.para, Math.max(0, tc.blocks.length - 1)))
    const para = tc.blocks[paraIndex]
    if (!para || para.type !== 'paragraph') {
      return { sectionIndex, blockIndex, offset: 0, cell: { row: cell.row, cell: cellIndex, para: 0 } }
    }
    const offset = Math.max(0, Math.min(pos.offset, paragraphLength(para)))
    return {
      sectionIndex,
      blockIndex,
      offset,
      cell: { row: Math.max(0, Math.min(cell.row, block.rows.length - 1)), cell: cellIndex, para: paraIndex },
    }
  }
  if (block.type !== 'paragraph') {
    return { sectionIndex, blockIndex, offset: 0 }
  }
  const offset = Math.max(0, Math.min(pos.offset, paragraphLength(block)))
  return { sectionIndex, blockIndex, offset }
}

export function documentCharCount(doc: Document): number {
  let n = 0
  for (const section of doc.sections) {
    for (const block of section.blocks) {
      if (block.type === 'paragraph') n += paragraphLength(block)
    }
  }
  return n
}

export interface RunHit {
  runIndex: number
  run: Run
  /** Offset within this run's contribution */
  offsetInRun: number
}

export function hitTestRun(paragraph: Paragraph, offset: number): RunHit {
  let cursor = 0
  for (let i = 0; i < paragraph.runs.length; i++) {
    const run = paragraph.runs[i]
    if (!run) continue
    const len = inlineToChar(run).length
    if (offset < cursor + len || (offset === cursor + len && i === paragraph.runs.length - 1)) {
      return { runIndex: i, run, offsetInRun: offset - cursor }
    }
    cursor += len
  }
  const lastIndex = Math.max(0, paragraph.runs.length - 1)
  const last = paragraph.runs[lastIndex]
  if (!last) {
    throw new AlmadocxError('empty_paragraph', 'paragraph has no runs')
  }
  return { runIndex: lastIndex, run: last, offsetInRun: inlineToChar(last).length }
}
