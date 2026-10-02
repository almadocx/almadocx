import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createApproximateMeasurer,
  extractPlainRange,
  layoutDocument,
  loadDocument,
  moveTableCellTab,
  moveTableCellVertical,
  paragraphPlainText,
  patchLayoutCellParagraph,
} from '../src/index.js'

const root = join(dirname(fileURLToPath(import.meta.url)), '../../..')
const fixture = join(root, 'fixtures/seamlesshr-golive-plan.docx')

describe('M3 table Word-parity (GoLive)', () => {
  it('Tab traverses cells left-to-right and wraps to next row', () => {
    const doc = loadDocument(new Uint8Array(readFileSync(fixture)), 'docx')
    // Find first multi-column table
    let bi = -1
    for (let i = 0; i < doc.sections[0]!.blocks.length; i++) {
      const b = doc.sections[0]!.blocks[i]!
      if (b.type === 'table' && (b.rows[0]?.cells.length ?? 0) >= 3) {
        bi = i
        break
      }
    }
    expect(bi).toBeGreaterThanOrEqual(0)
    const table = doc.sections[0]!.blocks[bi]!
    if (table.type !== 'table') return

    let pos = {
      sectionIndex: 0,
      blockIndex: bi,
      offset: 0,
      cell: { row: 0, cell: 0, para: 0 },
    }
    const cols = table.rows[0]!.cells.length
    for (let c = 1; c < cols; c++) {
      const r = moveTableCellTab(doc, pos, 1)
      expect(r.needsNewRow).toBeFalsy()
      expect(r.position.cell?.cell).toBe(c)
      expect(r.position.cell?.row).toBe(0)
      pos = r.position
    }
    const wrap = moveTableCellTab(doc, pos, 1)
    if (table.rows.length > 1) {
      expect(wrap.position.cell?.row).toBe(1)
      expect(wrap.position.cell?.cell).toBe(0)
    }
  })

  it('extractPlainRange returns text from a table cell', () => {
    const doc = loadDocument(new Uint8Array(readFileSync(fixture)), 'docx')
    let bi = -1
    for (let i = 0; i < doc.sections[0]!.blocks.length; i++) {
      const b = doc.sections[0]!.blocks[i]!
      if (b.type === 'table' && (b.rows[0]?.cells.length ?? 0) >= 2) {
        bi = i
        break
      }
    }
    expect(bi).toBeGreaterThanOrEqual(0)
    const table = doc.sections[0]!.blocks[bi]!
    if (table.type !== 'table') return
    const para = table.rows[0]!.cells[1]!.blocks[0]!
    const text = paragraphPlainText(para)
    expect(text.length).toBeGreaterThan(0)
    const plain = extractPlainRange(doc, {
      anchor: {
        sectionIndex: 0,
        blockIndex: bi,
        offset: 0,
        cell: { row: 0, cell: 1, para: 0 },
      },
      focus: {
        sectionIndex: 0,
        blockIndex: bi,
        offset: text.length,
        cell: { row: 0, cell: 1, para: 0 },
      },
    })
    expect(plain).toBe(text)
  })

  it('vertical move prefers same column', () => {
    const doc = loadDocument(new Uint8Array(readFileSync(fixture)), 'docx')
    let bi = -1
    for (let i = 0; i < doc.sections[0]!.blocks.length; i++) {
      const b = doc.sections[0]!.blocks[i]!
      if (b.type === 'table' && b.rows.length >= 2 && (b.rows[0]?.cells.length ?? 0) >= 2) {
        bi = i
        break
      }
    }
    expect(bi).toBeGreaterThanOrEqual(0)
    const next = moveTableCellVertical(
      doc,
      { sectionIndex: 0, blockIndex: bi, offset: 0, cell: { row: 0, cell: 1, para: 0 } },
      1,
      0,
    )
    expect(next.cell?.row).toBe(1)
    expect(next.cell?.cell).toBe(1)
  })

  it('Tab at last cell needsNewRow; insertRow then lands in new row', () => {
    let doc = loadDocument(new Uint8Array(readFileSync(fixture)), 'docx')
    let bi = -1
    for (let i = 0; i < doc.sections[0]!.blocks.length; i++) {
      const b = doc.sections[0]!.blocks[i]!
      if (b.type === 'table' && (b.rows[0]?.cells.length ?? 0) >= 2) {
        bi = i
        break
      }
    }
    expect(bi).toBeGreaterThanOrEqual(0)
    const table = doc.sections[0]!.blocks[bi]!
    if (table.type !== 'table') return
    const lastRow = table.rows.length - 1
    const lastCol = table.rows[lastRow]!.cells.length - 1
    const tab = moveTableCellTab(
      doc,
      {
        sectionIndex: 0,
        blockIndex: bi,
        offset: 0,
        cell: { row: lastRow, cell: lastCol, para: 0 },
      },
      1,
    )
    expect(tab.needsNewRow).toBe(true)
    const beforeRows = table.rows.length
    doc = applyOp(doc, {
      type: 'insertRow',
      sectionIndex: 0,
      blockIndex: bi,
      afterRow: lastRow,
    }).doc
    const next = doc.sections[0]!.blocks[bi]!
    expect(next.type).toBe('table')
    if (next.type === 'table') expect(next.rows.length).toBe(beforeRows + 1)
  })

  it('patchLayoutCellParagraph keeps cell x stable while typing', () => {
    let doc = loadDocument(new Uint8Array(readFileSync(fixture)), 'docx')
    let bi = -1
    for (let i = 0; i < doc.sections[0]!.blocks.length; i++) {
      const b = doc.sections[0]!.blocks[i]!
      if (b.type === 'table' && (b.rows[0]?.cells.length ?? 0) >= 3) {
        bi = i
        break
      }
    }
    expect(bi).toBeGreaterThanOrEqual(0)
    const measurer = createApproximateMeasurer()
    let layout = layoutDocument(doc, { measurer })
    const cellPath = { row: 0, cell: 1, para: 0 }
    const laid = layout.pages
      .flatMap((p) => p.blocks)
      .find((b) => b.kind === 'table' && b.blockIndex === bi)
    expect(laid?.kind).toBe('table')
    if (laid?.kind !== 'table') return
    const cell0 = laid.cells.find((c) => c.rowIndex === 0 && c.cellIndex === 1)
    expect(cell0).toBeTruthy()
    const x0 = cell0!.x

    doc = applyOp(doc, {
      type: 'insertText',
      position: {
        sectionIndex: 0,
        blockIndex: bi,
        offset: 0,
        cell: cellPath,
      },
      text: 'xyz',
    }).doc
    const patched = patchLayoutCellParagraph(doc, layout, 0, bi, cellPath, { measurer })
    layout = patched.layout
    const laid2 = layout.pages
      .flatMap((p) => p.blocks)
      .find((b) => b.kind === 'table' && b.blockIndex === bi)
    if (laid2?.kind !== 'table') return
    const cell1 = laid2.cells.find((c) => c.rowIndex === 0 && c.cellIndex === 1)
    expect(cell1!.x).toBeCloseTo(x0, 0)
  })
})
