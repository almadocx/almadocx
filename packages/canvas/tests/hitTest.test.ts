import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  createApproximateMeasurer,
  createEmptyDocument,
  layoutDocument,
  loadDocument,
  type LayoutPage,
  type LayoutParagraph,
  type LayoutResult,
  type Table,
} from '@almadocx/core'
import {
  hitTestPoint,
  pageEditableParagraphs,
  wordBounds,
} from '../src/selection/hitTest.js'

const root = join(dirname(fileURLToPath(import.meta.url)), '../../..')
const goliveFixture = join(root, 'fixtures/seamlesshr-golive-plan.docx')

function para(
  overrides: Partial<LayoutParagraph> & { blockIndex: number },
): LayoutParagraph {
  return {
    kind: 'paragraph',
    sectionIndex: 0,
    x: 10,
    y: 10,
    width: 100,
    height: 20,
    lines: [
      {
        y: 10,
        height: 20,
        baseline: 16,
        startOffset: 0,
        endOffset: 4,
        runs: [
          { text: 'abcd', x: 10, width: 40, font: '12px', color: '#000', startOffset: 0 },
        ],
      },
    ],
    ...overrides,
  }
}

describe('pageEditableParagraphs', () => {
  it('flattens body and table cell paragraphs; falls back to page.paragraphs', () => {
    const body = para({ blockIndex: 0 })
    const cellPara = para({ blockIndex: 1, cell: { row: 0, cell: 0, para: 0 }, y: 50 })
    const page: LayoutPage = {
      index: 0,
      width: 200,
      height: 200,
      blocks: [
        body,
        {
          kind: 'table',
          blockIndex: 1,
          sectionIndex: 0,
          x: 0,
          y: 40,
          width: 100,
          height: 40,
          cells: [
            {
              x: 0,
              y: 40,
              width: 100,
              height: 40,
              paragraphs: [cellPara],
              borderColor: '#000',
              rowIndex: 0,
              cellIndex: 0,
            },
          ],
          startRow: 0,
          endRow: 0,
        },
      ],
      paragraphs: [body],
    }
    expect(pageEditableParagraphs(page).length).toBe(2)

    const legacy: LayoutPage = {
      index: 0,
      width: 100,
      height: 100,
      blocks: undefined as unknown as LayoutPage['blocks'],
      paragraphs: [body],
    }
    expect(pageEditableParagraphs(legacy)).toEqual([body])
  })
})

describe('hitTestPoint tables', () => {
  it('targets the correct column in a synthetic multi-column table', () => {
    const table: Table = {
      id: 't1',
      type: 'table',
      props: {},
      gridCols: [2000, 3000, 2500],
      rows: [
        {
          id: 'r1',
          props: {},
          cells: ['A', 'B', 'C'].map((label, i) => ({
            id: `c${i}`,
            props: {},
            blocks: [
              {
                id: `p${i}`,
                type: 'paragraph' as const,
                props: {},
                runs: [
                  {
                    id: `run${i}`,
                    props: {},
                    content: { type: 'text' as const, text: label.repeat(4) },
                  },
                ],
              },
            ],
          })),
        },
      ],
    }

    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [table]
    const layout = layoutDocument(doc, { measurer: createApproximateMeasurer() })
    const laid = layout.pages[0]!.blocks.find((b) => b.kind === 'table')
    expect(laid?.kind).toBe('table')
    if (laid?.kind !== 'table') return

    const zoom = 1
    const pageGap = 24
    const viewportWidth = 900
    const page = layout.pages[0]!
    const marginX = Math.max(0, (viewportWidth - page.width * zoom) / 2)
    const pageTop = 24
    const pageLeft = marginX

    for (let i = 0; i < laid.cells.length; i++) {
      const cell = laid.cells[i]!
      const x = pageLeft + cell.x + cell.width / 2
      const y = pageTop + cell.y + 8
      const hit = hitTestPoint(layout, x, y, zoom, pageGap, 1, viewportWidth)
      expect(hit?.position.cell?.cell).toBe(i)
    }
  })

  it('targets each header column in the GoLive fixture (not always the last column)', () => {
    const bytes = new Uint8Array(readFileSync(goliveFixture))
    const doc = loadDocument(bytes, 'docx')
    const layout = layoutDocument(doc, { measurer: createApproximateMeasurer() })

    let pageIndex = -1
    let table: Extract<(typeof layout.pages)[number]['blocks'][number], { kind: 'table' }> | undefined
    for (let i = 0; i < layout.pages.length; i++) {
      for (const block of layout.pages[i]!.blocks) {
        if (block.kind !== 'table') continue
        const headerCells = block.cells.filter((c) => c.rowIndex === 0)
        if (headerCells.length >= 4) {
          pageIndex = i
          table = block
          break
        }
      }
      if (table) break
    }
    expect(pageIndex).toBeGreaterThanOrEqual(0)
    expect(table).toBeDefined()
    if (!table) return

    const page = layout.pages[pageIndex]!
    const zoom = 1
    const pageGap = 24
    const viewportWidth = 1200
    const marginX = Math.max(0, (viewportWidth - page.width * zoom) / 2)
    const pageTop = 24 + pageIndex * (page.height + pageGap)
    const pageLeft = marginX

    const headerCells = table.cells.filter((c) => c.rowIndex === 0).slice(0, 4)
    expect(headerCells.length).toBeGreaterThanOrEqual(4)

    for (const cell of headerCells) {
      const hit = hitTestPoint(
        layout,
        pageLeft + cell.x + cell.width / 2,
        pageTop + cell.y + 8,
        zoom,
        pageGap,
        1,
        viewportWidth,
      )
      expect(hit?.position.cell?.cell).toBe(cell.cellIndex)
    }
  })
})

describe('hitTestPoint body and edge cases', () => {
  it('returns undefined for empty layout and picks nearest page outside bounds', () => {
    expect(hitTestPoint({ pages: [] }, 0, 0, 1)).toBeUndefined()

    const p0 = para({ blockIndex: 0, y: 20 })
    const p1 = para({ blockIndex: 0, y: 20 })
    const layout: LayoutResult = {
      pages: [
        { index: 0, width: 100, height: 100, blocks: [p0], paragraphs: [p0] },
        { index: 1, width: 100, height: 100, blocks: [p1], paragraphs: [p1] },
      ],
    }
    const far = hitTestPoint(layout, 50, 5000, 1, 24, 1, 200)
    expect(far?.pageIndex).toBe(1)

    const inside = hitTestPoint(layout, 100, 50, 1, 24, 1, 200)
    expect(inside?.pageIndex).toBe(0)
    expect(inside?.position.offset).toBeGreaterThanOrEqual(0)
  })

  it('hit-tests body paragraphs to the right of runs and empty-line paragraphs', () => {
    const withRuns = para({ blockIndex: 0 })
    const emptyLines = para({
      blockIndex: 1,
      y: 80,
      lines: [],
      cell: { row: 0, cell: 0, para: 0 },
    })
    const page: LayoutPage = {
      index: 0,
      width: 300,
      height: 300,
      blocks: [withRuns, emptyLines],
      paragraphs: [withRuns, emptyLines],
    }
    const layout: LayoutResult = { pages: [page] }

    const endHit = hitTestPoint(layout, 200, 40, 1, 24, 1, 400)
    expect(endHit?.position.offset).toBe(4)

    const emptyHit = hitTestPoint(layout, 60, 90, 1, 24, 1, 400)
    expect(emptyHit?.position.blockIndex).toBe(1)
    expect(emptyHit?.position.cell).toEqual({ row: 0, cell: 0, para: 0 })
  })

  it('prefers table cells including empty-paragraph fallback and outside-cell scores', () => {
    const cellPara = para({
      blockIndex: 0,
      cell: { row: 0, cell: 0, para: 0 },
      x: 20,
      y: 20,
      width: 40,
    })
    const emptyCellPage: LayoutPage = {
      index: 0,
      width: 200,
      height: 200,
      blocks: [
        {
          kind: 'table',
          blockIndex: 0,
          sectionIndex: 0,
          x: 10,
          y: 10,
          width: 80,
          height: 60,
          cells: [
            {
              x: 10,
              y: 10,
              width: 40,
              height: 40,
              paragraphs: [cellPara],
              borderColor: '#000',
              rowIndex: 0,
              cellIndex: 0,
            },
            {
              x: 50,
              y: 10,
              width: 40,
              height: 40,
              paragraphs: [],
              borderColor: '#000',
              rowIndex: 0,
              cellIndex: 1,
            },
          ],
          startRow: 0,
          endRow: 0,
        },
      ],
      paragraphs: [],
    }
    const layout: LayoutResult = { pages: [emptyCellPage] }
    // viewportWidth 0 → no horizontal centering; pageTop = PAGE_TOP_PAD (24)
    const inPad = hitTestPoint(layout, 30, 24 + 30, 1, 24, 1, 0)
    expect(inPad?.position.cell).toBeDefined()

    const emptyCell = hitTestPoint(layout, 70, 24 + 30, 1, 24, 1, 0)
    expect(emptyCell?.position.cell?.cell).toBe(1)
    expect(emptyCell?.position.offset).toBe(0)

    // Outside table entirely → body default
    const outside = hitTestPoint(layout, 180, 24 + 180, 1, 24, 1, 0)
    expect(outside?.position).toEqual({ sectionIndex: 0, blockIndex: 0, offset: 0 })
  })

  it('handles zero-width runs and multi-line y selection', () => {
    const multi = para({
      blockIndex: 0,
      height: 60,
      lines: [
        {
          y: 10,
          height: 20,
          baseline: 16,
          startOffset: 0,
          endOffset: 2,
          runs: [{ text: 'ab', x: 10, width: 0, font: '12px', color: '#000', startOffset: 0 }],
        },
        {
          y: 30,
          height: 20,
          baseline: 16,
          startOffset: 2,
          endOffset: 4,
          runs: [{ text: 'cd', x: 10, width: 20, font: '12px', color: '#000', startOffset: 2 }],
        },
      ],
    })
    const layout: LayoutResult = {
      pages: [{ index: 0, width: 200, height: 200, blocks: [multi], paragraphs: [multi] }],
    }
    const hit = hitTestPoint(layout, 15, 55, 1, 24, 1, 400)
    expect(hit?.position.offset).toBeGreaterThanOrEqual(2)

    const zero = hitTestPoint(layout, 12, 25, 1, 24, 1, 400)
    expect(zero?.position.offset).toBe(0)
  })
})

describe('wordBounds', () => {
  it('covers empty, word, boundary, and punctuation branches', () => {
    expect(wordBounds('', 0)).toEqual({ start: 0, end: 0 })
    expect(wordBounds('hello', 2)).toEqual({ start: 0, end: 5 })
    expect(wordBounds('hello world', 5)).toEqual({ start: 0, end: 5 })
    expect(wordBounds('hello world', 11)).toEqual({ start: 6, end: 11 })
    expect(wordBounds('a,,,b', 2)).toEqual({ start: 1, end: 4 })
    expect(wordBounds('a\nb', 1)).toEqual({ start: 0, end: 1 })
    expect(wordBounds('x', 0)).toEqual({ start: 0, end: 1 })
    expect(wordBounds('!!!', 1).end).toBeGreaterThan(wordBounds('!!!', 1).start)
    expect(wordBounds(' \n ', 1)).toEqual({ start: 0, end: 1 })
    // Lone newline: punctuation branch collapses start===end then expands ±1
    expect(wordBounds('\n', 0)).toEqual({ start: 0, end: 1 })
    expect(wordBounds('\n\n', 1)).toEqual({ start: 0, end: 2 })
    // Offset past end on a single punctuation char
    expect(wordBounds('!', 1)).toEqual({ start: 0, end: 1 })
    // Word-before branch (cursor on punctuation after a word)
    expect(wordBounds('ab!', 2)).toEqual({ start: 0, end: 2 })
    expect(wordBounds('hello', 5)).toEqual({ start: 0, end: 5 })
  })
})

describe('hitTestInPage edge branches', () => {
  it('uses paragraphs fallback when blocks are missing and scores cells above/below', () => {
    const body = para({ blockIndex: 0, y: 50, x: 10, width: 80 })
    const legacy: LayoutPage = {
      index: 0,
      width: 200,
      height: 200,
      blocks: undefined as unknown as LayoutPage['blocks'],
      paragraphs: [body],
    }
    const layoutLegacy: LayoutResult = { pages: [legacy] }
    const hit = hitTestPoint(layoutLegacy, 50, 24 + 60, 1, 24, 1, 0)
    expect(hit?.position.blockIndex).toBe(0)

    const cellPara = para({
      blockIndex: 0,
      cell: { row: 0, cell: 0, para: 0 },
      x: 20,
      y: 40,
      width: 40,
      height: 20,
    })
    const tablePage: LayoutPage = {
      index: 0,
      width: 200,
      height: 200,
      blocks: [
        {
          kind: 'table',
          blockIndex: 0,
          sectionIndex: 0,
          x: 10,
          y: 30,
          width: 80,
          height: 50,
          cells: [
            {
              x: 20,
              y: 40,
              width: 40,
              height: 20,
              paragraphs: [cellPara],
              borderColor: '#000',
              rowIndex: 0,
              cellIndex: 0,
            },
          ],
          startRow: 0,
          endRow: 0,
        },
      ],
      paragraphs: [],
    }
    const layoutTable: LayoutResult = { pages: [tablePage] }
    // Above the cell but inside table padding band
    const above = hitTestPoint(layoutTable, 40, 24 + 32, 1, 24, 1, 0)
    expect(above?.position.cell?.cell).toBe(0)
    // Below the cell
    const below = hitTestPoint(layoutTable, 40, 24 + 70, 1, 24, 1, 0)
    expect(below?.position.cell?.cell).toBe(0)
  })

  it('hit-tests empty-text runs and empty-line body fallbacks', () => {
    const emptyText = para({
      blockIndex: 0,
      lines: [
        {
          y: 10,
          height: 20,
          baseline: 16,
          startOffset: 0,
          endOffset: 1,
          runs: [{ text: '', x: 10, width: 10, font: '12px', color: '#000', startOffset: 0 }],
        },
      ],
    })
    const layout: LayoutResult = {
      pages: [
        { index: 0, width: 200, height: 200, blocks: [emptyText], paragraphs: [emptyText] },
      ],
    }
    const hit = hitTestPoint(layout, 15, 24 + 15, 1, 24, 1, 0)
    expect(hit?.position.offset).toBeGreaterThanOrEqual(0)
  })
})
