import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  createApproximateMeasurer,
  layoutDocument,
  loadDocument,
  nextTabStopTwips,
  resolveParagraphProps,
} from '../src/index.js'

const root = join(dirname(fileURLToPath(import.meta.url)), '../../..')
const fixture = join(root, 'fixtures/iou-form.docx')

describe('IOU / Imprest form fixture', () => {
  it('parses VML pict logo, table merges, and Footer tab stops', () => {
    const bytes = new Uint8Array(readFileSync(fixture))
    const doc = loadDocument(bytes, 'docx')
    const section = doc.sections[0]!

    const headerPara = section.blocks[0]
    expect(headerPara?.type).toBe('paragraph')
    if (headerPara?.type !== 'paragraph') return
    const imageRun = headerPara.runs.find((r) => r.content.type === 'image')
    expect(imageRun?.content.type).toBe('image')
    if (imageRun?.content.type !== 'image') return
    const media = doc.media[imageRun.content.mediaId]
    expect(media?.bytes.byteLength).toBeGreaterThan(10_000)

    const table = section.blocks.find((b) => b.type === 'table')
    expect(table?.type).toBe('table')
    if (table?.type !== 'table') return
    expect(table.props.indentTwips).toBe(108)
    expect(table.rows[1]?.cells[0]?.props.vMerge).toBe('restart')
    expect(table.rows[2]?.cells[0]?.props.vMerge).toBe('continue')
    expect(table.rows[3]?.cells[0]?.props.vMerge).toBe('restart')

    expect(section.footer).toBeDefined()
    const footerPara = section.footer!.blocks[0]
    expect(footerPara?.type).toBe('paragraph')
    if (footerPara?.type !== 'paragraph') return
    const footerProps = resolveParagraphProps(doc, footerPara)
    expect(footerProps.tabs?.some((t) => t.alignment === 'right' && t.position === 9360)).toBe(true)
    expect(footerPara.runs.some((r) => r.content.type === 'text' && r.content.text.includes('FIN-'))).toBe(
      true,
    )
  })

  it('wraps overflow tabs onto the next line (form underline rows)', () => {
    // Past the right edge, nextTabStopTwips signals wrap by returning beyond contentWidth.
    const past = nextTabStopTwips(9360, undefined, 9360)
    expect(past.position).toBeGreaterThan(9360)
  })

  it('layouts to multiple pages without mid-page Y resets', () => {
    const bytes = new Uint8Array(readFileSync(fixture))
    const doc = loadDocument(bytes, 'docx')
    const layout = layoutDocument(doc, { measurer: createApproximateMeasurer() })
    expect(layout.pages.length).toBeGreaterThanOrEqual(2)

    for (const page of layout.pages) {
      let prevBottom = -1
      for (const block of page.blocks) {
        if (prevBottom >= 0 && block.y + 5 < prevBottom) {
          expect.fail(`Y reset on page: block at ${block.y} after ${prevBottom}`)
        }
        prevBottom = Math.max(prevBottom, block.y + block.height)
      }
      expect(page.footer?.paragraphs.length).toBeGreaterThan(0)
    }
  })
})
