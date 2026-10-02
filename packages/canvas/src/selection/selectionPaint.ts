import type { DocRange, LayoutPage, LayoutParagraph, LayoutResult } from '@almadocx/core'
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
