import { describe, expect, it } from 'vitest'
import type { LayoutResult } from '@almadocx/core'
import {
  computePageLayoutMetrics,
  contentMarginX,
  pageOriginInView,
  PAGE_GAP_DEFAULT,
} from '../src/render/pageLayout.js'

function layout(pages: Array<{ w: number; h: number }>): LayoutResult {
  return {
    pages: pages.map((p, i) => ({
      index: i,
      width: p.w,
      height: p.h,
      blocks: [],
      paragraphs: [],
    })),
  }
}

describe('pageLayout', () => {
  it('computes metrics for multi-column stacks', () => {
    const metrics = computePageLayoutMetrics(layout([{ w: 100, h: 200 }, { w: 80, h: 180 }]), 2, 10, 2)
    expect(metrics.cols).toBe(2)
    expect(metrics.pageW).toBe(100)
    expect(metrics.pageH).toBe(200)
    expect(metrics.contentWidth).toBeGreaterThan(0)
    expect(metrics.contentHeight).toBeGreaterThan(0)

    const single = computePageLayoutMetrics(layout([]), 1, PAGE_GAP_DEFAULT, 0)
    expect(single.cols).toBe(1)
    expect(single.pageW).toBe(0)
  })

  it('centers content when viewport is wider, else returns 0 margin', () => {
    expect(contentMarginX(200, 0)).toBe(0)
    expect(contentMarginX(200, -1)).toBe(0)
    expect(contentMarginX(200, 400)).toBe(100)
    expect(contentMarginX(400, 200)).toBe(0)
  })

  it('places page origins by row/column', () => {
    const metrics = computePageLayoutMetrics(
      layout([
        { w: 100, h: 100 },
        { w: 100, h: 100 },
        { w: 100, h: 100 },
      ]),
      1,
      10,
      2,
    )
    const p0 = pageOriginInView(metrics, 0, 1, 10, 500)
    const p1 = pageOriginInView(metrics, 1, 1, 10, 500)
    const p2 = pageOriginInView(metrics, 2, 1, 10, 500)
    expect(p1.left).toBeGreaterThan(p0.left)
    expect(p1.top).toBe(p0.top)
    expect(p2.top).toBeGreaterThan(p0.top)
    expect(p2.left).toBe(p0.left)
  })
})
