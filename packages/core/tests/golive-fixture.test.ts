import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { createApproximateMeasurer, layoutDocument, loadDocument } from '../src/index.js'

const root = join(dirname(fileURLToPath(import.meta.url)), '../../..')
const fixture = join(root, 'fixtures/seamlesshr-golive-plan.docx')

describe('SeamlessHR GoLive Plan fixture', () => {
  it('parses section margins from sectPr', () => {
    const bytes = new Uint8Array(readFileSync(fixture))
    const doc = loadDocument(bytes, 'docx')
    expect(doc.sections[0]!.properties.margins.left).toBe(1800)
    expect(doc.sections[0]!.properties.margins.right).toBe(1800)
    expect(doc.sections[0]!.footer).toBeDefined()
  })

  it('layouts without overlapping body blocks or mid-page Y resets', () => {
    const bytes = new Uint8Array(readFileSync(fixture))
    const doc = loadDocument(bytes, 'docx')
    const layout = layoutDocument(doc, { measurer: createApproximateMeasurer() })
    expect(layout.pages.length).toBeGreaterThan(10)

    let overlaps = 0
    let yResets = 0
    for (const page of layout.pages) {
      let prevBottom = -1
      for (const block of page.blocks) {
        if (prevBottom >= 0 && block.y + 5 < prevBottom) yResets += 1
        prevBottom = Math.max(prevBottom, block.y + block.height)
      }
      for (let i = 0; i < page.blocks.length; i++) {
        for (let j = i + 1; j < page.blocks.length; j++) {
          const a = page.blocks[i]!
          const b = page.blocks[j]!
          const yOverlap = a.y < b.y + b.height && b.y < a.y + a.height
          const xOverlap = a.x < b.x + b.width && b.x < a.x + a.width
          const depth = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y)
          if (yOverlap && xOverlap && depth > 4) overlaps += 1
        }
      }
      // No invented header page numbers when the doc only has a footer
      expect(page.header?.pageNumberText).toBeUndefined()
    }
    expect(overlaps).toBe(0)
    expect(yResets).toBe(0)
  })
})
