import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  loadDocument,
  paragraphPlainText,
  saveDocument,
} from '../src/index.js'

function sampleDoc() {
  let doc = createEmptyDocument('docx')
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
    text: 'Almadocx Phase 1',
  }).doc
  doc = applyOp(doc, {
    type: 'splitParagraph',
    position: { sectionIndex: 0, blockIndex: 0, offset: 9 },
  }).doc
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
    text: 'Second paragraph with ',
  }).doc
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 1, offset: 22 },
    text: 'bold',
    props: { bold: true },
  }).doc
  return doc
}

describe('DOCX/ODT round-trip', () => {
  it('round-trips DOCX text', () => {
    const doc = sampleDoc()
    const bytes = saveDocument(doc, 'docx')
    const loaded = loadDocument(bytes, 'docx')
    expect(paragraphPlainText(loaded.sections[0]!.blocks[0]!)).toContain('Almadocx')
    expect(loaded.sections[0]!.blocks.length).toBeGreaterThanOrEqual(2)
  })

  it('round-trips ODT text', () => {
    const doc = { ...sampleDoc(), package: { ...sampleDoc().package, sourceFormat: 'odt' as const } }
    const bytes = saveDocument(doc, 'odt')
    const loaded = loadDocument(bytes, 'odt')
    const texts = loaded.sections[0]!.blocks.map((b) => paragraphPlainText(b))
    expect(texts.join(' ')).toContain('Almadocx')
  })
})
