import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  createApproximateMeasurer,
  createEmptyDocument,
  layoutDocument,
  loadDocument,
  type Table,
} from '@almadocx/core'
import { hitTestPoint } from '../src/selection/hitTest.js'

const root = join(dirname(fileURLToPath(import.meta.url)), '../../..')
const goliveFixture = join(root, 'fixtures/seamlesshr-golive-plan.docx')

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

    const page = layout.pages[21]!
    const table = page.blocks.find((b) => b.kind === 'table' && b.cells.length >= 4)
    expect(table?.kind).toBe('table')
    if (table?.kind !== 'table') return

    const zoom = 1
    const pageGap = 24
    const viewportWidth = 1200
    const marginX = Math.max(0, (viewportWidth - page.width * zoom) / 2)
    const pageTop = 24 + 21 * (page.height + pageGap)
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
