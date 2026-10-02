import type { LayoutResult } from '@almadocx/core'

/** Vertical breathing room above first page / between page rows. */
export const PAGE_TOP_PAD = 24
export const PAGE_GAP_DEFAULT = 24

export interface PageLayoutMetrics {
  cols: number
  pageW: number
  pageH: number
  /** Exact width of the page stack (no viewport padding). */
  contentWidth: number
  contentHeight: number
}

export function computePageLayoutMetrics(
  layout: LayoutResult,
  zoom: number,
  pageGap: number,
  pageColumns: number,
): PageLayoutMetrics {
  const cols = Math.max(1, pageColumns)
  const pageW = Math.max(...layout.pages.map((p) => p.width), 0)
  const pageH = Math.max(...layout.pages.map((p) => p.height), 0)
  const rows = Math.ceil(Math.max(layout.pages.length, 1) / cols)
  const contentWidth = cols * pageW * zoom + pageGap * Math.max(0, cols - 1) * zoom
  const contentHeight =
    PAGE_TOP_PAD * zoom +
    rows * pageH * zoom +
    pageGap * Math.max(0, rows - 1) * zoom +
    PAGE_TOP_PAD * zoom
  return { cols, pageW, pageH, contentWidth, contentHeight }
}

/**
 * Horizontal inset so the page stack is centered in the viewport when it fits,
 * or left-aligned (scrollable) when wider than the viewport.
 */
export function contentMarginX(contentWidth: number, viewportWidth: number): number {
  if (viewportWidth <= 0) return 0
  return Math.max(0, (viewportWidth - contentWidth) / 2)
}

export function pageOriginInView(
  metrics: PageLayoutMetrics,
  pageIndex: number,
  zoom: number,
  pageGap: number,
  viewportWidth: number,
): { left: number; top: number } {
  const { cols, pageW, pageH, contentWidth } = metrics
  const row = Math.floor(pageIndex / cols)
  const col = pageIndex % cols
  const marginX = contentMarginX(contentWidth, viewportWidth)
  const pageTop = PAGE_TOP_PAD * zoom + row * (pageH * zoom + pageGap * zoom)
  const pageLeft = marginX + col * (pageW * zoom + pageGap * zoom)
  return { left: pageLeft, top: pageTop }
}
