import type {
  DocRange,
  LayoutPage,
  LayoutParagraph,
  LayoutResult,
  LayoutTable,
  LayoutTableCell,
} from '@almadocx/core'
import { normalizeRange } from '@almadocx/core'
import { pageEditableParagraphs } from './hitTest.js'
import {
  computePageLayoutMetrics,
  pageOriginInView,
  PAGE_GAP_DEFAULT,
} from '../render/pageLayout.js'

type PosLike = {
  sectionIndex: number
  blockIndex: number
  offset: number
  cell?: { row: number; cell: number; para: number }
}

function sameCell(
  a: { row: number; cell: number; para: number } | undefined,
  b: { row: number; cell: number; para: number } | undefined,
): boolean {
  return (!a && !b) || (!!a && !!b && a.row === b.row && a.cell === b.cell && a.para === b.para)
}

function cmpCell(
  a: { row: number; cell: number; para: number } | undefined,
  b: { row: number; cell: number; para: number } | undefined,
): number {
  if (!a && !b) return 0
  if (!a) return -1
  if (!b) return 1
  if (a.row !== b.row) return a.row - b.row
  if (a.cell !== b.cell) return a.cell - b.cell
  return a.para - b.para
}

/**
 * True when the selection spans multiple cells of a single table block
 * (Word-style rectangular cell selection) rather than text inside one cell.
 */
export function isRectCellSelection(start: PosLike, end: PosLike): boolean {
  return (
    !!start.cell &&
    !!end.cell &&
    start.sectionIndex === end.sectionIndex &&
    start.blockIndex === end.blockIndex &&
    (start.cell.row !== end.cell.row || start.cell.cell !== end.cell.cell)
  )
}

/** Layout cells of `table` inside the inclusive rectangle [r0..r1] × [c0..c1]. */
export function cellsInRect(
  table: LayoutTable,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): LayoutTableCell[] {
  const rowLo = Math.min(r0, r1)
  const rowHi = Math.max(r0, r1)
  const colLo = Math.min(c0, c1)
  const colHi = Math.max(c0, c1)
  return table.cells.filter(
    (cell) =>
      cell.rowIndex >= rowLo &&
      cell.rowIndex <= rowHi &&
      cell.cellIndex >= colLo &&
      cell.cellIndex <= colHi,
  )
}

/** True if paragraph participates in [start, end] selection. */
export function paragraphIntersectsSelection(
  para: LayoutParagraph,
  start: PosLike,
  end: PosLike,
): boolean {
  if (para.sectionIndex < start.sectionIndex || para.sectionIndex > end.sectionIndex) {
    return false
  }
  if (para.sectionIndex === start.sectionIndex && para.blockIndex < start.blockIndex) {
    return false
  }
  if (para.sectionIndex === end.sectionIndex && para.blockIndex > end.blockIndex) {
    return false
  }
  if (
    para.sectionIndex === start.sectionIndex &&
    para.blockIndex === start.blockIndex &&
    (para.cell || start.cell || end.cell)
  ) {
    if (cmpCell(para.cell, start.cell) < 0) return false
    if (cmpCell(para.cell, end.cell) > 0) return false
  }
  return true
}

export function lineSelectionSpan(
  para: LayoutParagraph,
  line: LayoutParagraph['lines'][number],
  start: PosLike,
  end: PosLike,
): { selStart: number; selEnd: number } | undefined {
  const isStartPara =
    para.sectionIndex === start.sectionIndex &&
    para.blockIndex === start.blockIndex &&
    sameCell(para.cell, start.cell)
  const isEndPara =
    para.sectionIndex === end.sectionIndex &&
    para.blockIndex === end.blockIndex &&
    sameCell(para.cell, end.cell)

  if (isStartPara && line.endOffset <= start.offset) return undefined
  if (isEndPara && line.startOffset >= end.offset) return undefined

  const selStart = isStartPara ? Math.max(line.startOffset, start.offset) : line.startOffset
  const selEnd = isEndPara ? Math.min(line.endOffset, end.offset) : line.endOffset
  if (selEnd <= selStart) return undefined
  return { selStart, selEnd }
}

export function offsetToX(line: LayoutParagraph['lines'][number], offset: number): number {
  if (line.runs.length === 0) return 0
  for (const run of line.runs) {
    const charLen = run.image ? 1 : run.text.length > 0 ? run.text.length : 1
    const runEnd = run.startOffset + charLen
    if (offset <= runEnd) {
      const local = offset - run.startOffset
      if (local <= 0) return run.x
      const ratio = Math.min(1, local / charLen)
      return run.x + run.width * ratio
    }
  }
  const last = line.runs[line.runs.length - 1]!
  return last.x + last.width
}

/** Map an X coordinate on a line to the nearest character offset. */
export function xToOffset(line: LayoutParagraph['lines'][number], x: number): number {
  if (line.runs.length === 0) return line.startOffset
  for (const run of line.runs) {
    if (x <= run.x + run.width) {
      const charLen = run.image ? 1 : run.text.length > 0 ? run.text.length : 1
      const ratio = run.width === 0 ? 0 : Math.max(0, Math.min(1, (x - run.x) / run.width))
      return run.startOffset + Math.round(ratio * charLen)
    }
  }
  return line.endOffset
}

function sameCellPath(
  a: { row: number; cell: number; para: number } | undefined,
  b: { row: number; cell: number; para: number } | undefined,
): boolean {
  return (!a && !b) || (!!a && !!b && a.row === b.row && a.cell === b.cell && a.para === b.para)
}

/**
 * Move caret up/down by visual layout line within the current paragraph.
 * Returns undefined when already on the first/last line (caller should leave the para).
 */
export function moveByLayoutLine(
  layout: LayoutResult,
  pos: {
    sectionIndex: number
    blockIndex: number
    offset: number
    cell?: { row: number; cell: number; para: number }
  },
  direction: -1 | 1,
  preferX?: number,
): { sectionIndex: number; blockIndex: number; offset: number; cell?: { row: number; cell: number; para: number } } | undefined {
  for (const page of layout.pages) {
    for (const para of pageEditableParagraphs(page)) {
      if (para.sectionIndex !== pos.sectionIndex || para.blockIndex !== pos.blockIndex) continue
      if (!sameCellPath(para.cell, pos.cell)) continue
      if (para.lines.length === 0) return undefined

      let lineIndex = 0
      for (let i = 0; i < para.lines.length; i++) {
        const line = para.lines[i]!
        const isLast = i === para.lines.length - 1
        if (
          pos.offset >= line.startOffset &&
          (pos.offset < line.endOffset || (pos.offset === line.endOffset && isLast))
        ) {
          lineIndex = i
          break
        }
        if (pos.offset >= line.startOffset) lineIndex = i
      }

      const nextIndex = lineIndex + direction
      if (nextIndex < 0 || nextIndex >= para.lines.length) return undefined

      const curLine = para.lines[lineIndex]!
      const nextLine = para.lines[nextIndex]!
      const x = preferX ?? offsetToX(curLine, pos.offset)
      const offset = xToOffset(nextLine, x)
      const out: {
        sectionIndex: number
        blockIndex: number
        offset: number
        cell?: { row: number; cell: number; para: number }
      } = {
        sectionIndex: para.sectionIndex,
        blockIndex: para.blockIndex,
        offset,
      }
      if (para.cell) out.cell = { ...para.cell }
      return out
    }
  }
  return undefined
}

export function paintSelection(
  ctx: CanvasRenderingContext2D,
  page: LayoutPage,
  selection: DocRange,
): void {
  const { start, end } = normalizeRange(selection)
  if (
    start.sectionIndex === end.sectionIndex &&
    start.blockIndex === end.blockIndex &&
    start.offset === end.offset &&
    sameCell(start.cell, end.cell)
  ) {
    return
  }
  ctx.fillStyle = 'rgba(47, 107, 255, 0.28)'

  // Word-style rectangular multi-cell selection: paint whole cells.
  if (isRectCellSelection(start, end)) {
    const r0 = start.cell!.row
    const c0 = start.cell!.cell
    const r1 = end.cell!.row
    const c1 = end.cell!.cell
    for (const block of page.blocks ?? []) {
      if (block.kind !== 'table') continue
      if (block.sectionIndex !== start.sectionIndex || block.blockIndex !== start.blockIndex) {
        continue
      }
      for (const cell of cellsInRect(block, r0, c0, r1, c1)) {
        ctx.fillRect(cell.x, cell.y, cell.width, cell.height)
      }
    }
    return
  }

  for (const para of pageEditableParagraphs(page)) {
    if (!paragraphIntersectsSelection(para, start, end)) continue
    for (const line of para.lines) {
      const span = lineSelectionSpan(para, line, start, end)
      if (!span) continue
      const x0 = offsetToX(line, span.selStart)
      const x1 = offsetToX(line, span.selEnd)
      ctx.fillRect(x0, line.y, Math.max(1, x1 - x0), line.height)
    }
  }
}

export interface CaretScreenRect {
  top: number
  bottom: number
  left: number
  pageIndex: number
}

export function getCaretScreenRect(
  layout: LayoutResult,
  selection: DocRange,
  zoom: number,
  pageGap = PAGE_GAP_DEFAULT,
  pageColumns = 1,
  viewportWidth = 0,
): CaretScreenRect | undefined {
  const pos = selection.focus
  const metrics = computePageLayoutMetrics(layout, zoom, pageGap, pageColumns)

  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]!
    const { left: pageLeft, top: pageTop } = pageOriginInView(
      metrics,
      i,
      zoom,
      pageGap,
      viewportWidth,
    )

    for (const para of pageEditableParagraphs(page)) {
      if (para.sectionIndex !== pos.sectionIndex || para.blockIndex !== pos.blockIndex) continue
      if (pos.cell || para.cell) {
        if (!sameCell(pos.cell, para.cell)) continue
      }
      for (let li = 0; li < para.lines.length; li++) {
        const line = para.lines[li]!
        const isLastLine = li === para.lines.length - 1
        const inLine =
          pos.offset >= line.startOffset &&
          (pos.offset < line.endOffset || (pos.offset === line.endOffset && isLastLine))
        if (!inLine) continue
        const x = offsetToX(line, pos.offset)
        const top = pageTop + line.y * zoom
        return {
          top,
          bottom: top + line.height * zoom,
          left: pageLeft + x * zoom,
          pageIndex: page.index,
        }
      }
    }
  }
  return undefined
}
