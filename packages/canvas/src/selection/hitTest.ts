import type { CellPath, DocPosition, LayoutPage, LayoutParagraph, LayoutResult } from '@almadocx/core'
import {
  computePageLayoutMetrics,
  pageOriginInView,
  PAGE_GAP_DEFAULT,
} from '../render/pageLayout.js'

export interface HitTestResult {
  position: DocPosition
  pageIndex: number
}

/** Flatten all editable paragraphs on a page (body + table cells). */
export function pageEditableParagraphs(page: LayoutPage): LayoutParagraph[] {
  const out: LayoutParagraph[] = []
  for (const block of page.blocks ?? page.paragraphs) {
    if (block.kind === 'paragraph') out.push(block)
    else if (block.kind === 'table') {
      for (const cell of block.cells) out.push(...cell.paragraphs)
    }
  }
  return out
}

export function hitTestPoint(
  layout: LayoutResult,
  x: number,
  y: number,
  zoom: number,
  pageGap = PAGE_GAP_DEFAULT,
  pageColumns = 1,
  viewportWidth = 0,
): HitTestResult | undefined {
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
    const pageBottom = pageTop + page.height * zoom
    const pageRight = pageLeft + page.width * zoom
    if (y >= pageTop && y <= pageBottom && x >= pageLeft && x <= pageRight) {
      const localX = (x - pageLeft) / zoom
      const localY = (y - pageTop) / zoom
      return { position: hitTestInPage(page, localX, localY), pageIndex: page.index }
    }
  }

  let bestI = 0
  let bestDist = Infinity
  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]!
    const { top: pageTop } = pageOriginInView(metrics, i, zoom, pageGap, viewportWidth)
    const mid = pageTop + (page.height * zoom) / 2
    const d = Math.abs(y - mid)
    if (d < bestDist) {
      bestDist = d
      bestI = i
    }
  }
  const last = layout.pages[bestI]
  if (!last) return undefined
  const { left: pageLeft, top: pageTop } = pageOriginInView(
    metrics,
    bestI,
    zoom,
    pageGap,
    viewportWidth,
  )
  return {
    position: hitTestInPage(last, (x - pageLeft) / zoom, (y - pageTop) / zoom),
    pageIndex: last.index,
  }
}

function hitTestInPage(page: LayoutPage, x: number, y: number): DocPosition {
  // Prefer a table under the point; pick nearest cell by distance (Word-like column targeting).
  let best:
    | {
        sectionIndex: number
        blockIndex: number
        cell: { x: number; y: number; width: number; height: number; rowIndex: number; cellIndex: number; paragraphs: LayoutParagraph[] }
        score: number
      }
    | undefined
  for (const block of page.blocks ?? []) {
    if (block.kind !== 'table') continue
    const inTable =
      x >= block.x - 4 &&
      x <= block.x + block.width + 4 &&
      y >= block.y - 4 &&
      y <= block.y + block.height + 4
    if (!inTable) continue
    for (const cell of block.cells) {
      const dx =
        x < cell.x ? cell.x - x : x > cell.x + cell.width ? x - (cell.x + cell.width) : 0
      const dy =
        y < cell.y ? cell.y - y : y > cell.y + cell.height ? y - (cell.y + cell.height) : 0
      const score = dy * 1000 + dx
      if (!best || score < best.score) {
        best = {
          sectionIndex: block.sectionIndex,
          blockIndex: block.blockIndex,
          cell,
          score,
        }
      }
    }
  }
  if (best) {
    return (
      hitTestInParagraphs(best.cell.paragraphs, x, y) ?? {
        sectionIndex: best.sectionIndex,
        blockIndex: best.blockIndex,
        offset: 0,
        cell: {
          row: best.cell.rowIndex,
          cell: best.cell.cellIndex,
          para: 0,
        },
      }
    )
  }

  // Body paragraphs: prefer one whose vertical band contains y; break ties by x
  const paras = (page.blocks ?? page.paragraphs).filter(
    (b): b is LayoutParagraph => b.kind === 'paragraph',
  )
  return hitTestInParagraphs(paras, x, y) ?? { sectionIndex: 0, blockIndex: 0, offset: 0 }
}

function hitTestInParagraphs(
  paras: LayoutParagraph[],
  x: number,
  y: number,
): DocPosition | undefined {
  if (paras.length === 0) return undefined

  // Table rows lay out every cell paragraph on the same baseline Y; prefer the
  // paragraph whose horizontal band actually contains x before Y-only heuristics.
  let best: LayoutParagraph | undefined
  let bestScore = Infinity
  for (const para of paras) {
    const dy =
      y < para.y ? para.y - y : y > para.y + para.height ? y - (para.y + para.height) : 0
    const dx = horizontalDist(para, x)
    const containsX = dx === 0
    const containsY = dy === 0
    const score =
      (containsX && containsY ? 0 : containsX ? 1 : containsY ? 2 : 3) * 1_000_000 +
      dy * 1000 +
      dx
    if (score < bestScore) {
      bestScore = score
      best = para
    }
  }
  // paras is non-empty, so the score loop always picks a best paragraph.
  const chosen = best!
  if (chosen.lines.length === 0) {
    const pos: DocPosition = {
      sectionIndex: chosen.sectionIndex,
      blockIndex: chosen.blockIndex,
      offset: 0,
    }
    if (chosen.cell) pos.cell = { ...chosen.cell }
    return pos
  }

  let line = chosen.lines[0]!
  for (const l of chosen.lines) {
    if (y >= l.y) line = l
  }
  for (const run of line.runs) {
    if (x <= run.x + run.width) {
      const ratio = run.width === 0 ? 0 : Math.max(0, Math.min(1, (x - run.x) / run.width))
      const local = Math.round(ratio * (run.text.length || 1))
      const pos: DocPosition = {
        sectionIndex: chosen.sectionIndex,
        blockIndex: chosen.blockIndex,
        offset: run.startOffset + local,
      }
      if (chosen.cell) pos.cell = { ...chosen.cell }
      return pos
    }
  }
  const pos: DocPosition = {
    sectionIndex: chosen.sectionIndex,
    blockIndex: chosen.blockIndex,
    offset: line.endOffset,
  }
  if (chosen.cell) pos.cell = { ...chosen.cell }
  return pos
}

function horizontalDist(para: LayoutParagraph, x: number): number {
  if (x < para.x) return para.x - x
  if (x > para.x + para.width) return x - (para.x + para.width)
  return 0
}

export function wordBounds(text: string, offset: number): { start: number; end: number } {
  if (text.length === 0) return { start: 0, end: 0 }
  const o = Math.max(0, Math.min(offset, text.length))
  const isWord = (ch: string) => /[\p{L}\p{N}_]/u.test(ch)
  let start = o
  let end = o
  if (o < text.length && isWord(text[o]!)) {
    while (start > 0 && isWord(text[start - 1]!)) start -= 1
    while (end < text.length && isWord(text[end]!)) end += 1
  } else if (o > 0 && isWord(text[o - 1]!)) {
    start = o - 1
    end = o
    while (start > 0 && isWord(text[start - 1]!)) start -= 1
    // end already sits on a non-word (or past the last char); no forward expand
  } else {
    while (start > 0 && !isWord(text[start - 1]!) && text[start - 1] !== '\n') start -= 1
    while (end < text.length && !isWord(text[end]!) && text[end] !== '\n') end += 1
    if (start === end) {
      start = Math.max(0, o - 1)
      end = Math.min(text.length, o + 1)
    }
  }
  return { start, end }
}

export type { CellPath }
