import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  layoutDocument,
  saveDocument,
  loadDocument,
} from '@almadocx/core'

describe('visual harness foundations', () => {
  it('layout is deterministic across runs', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'Deterministic layout check for Almadocx harness. '.repeat(12),
    }).doc
    const a = layoutDocument(doc)
    const b = layoutDocument(doc)
    expect(JSON.stringify(a)).toEqual(JSON.stringify(b))
    expect(a.pages[0]!.paragraphs[0]!.lines.length).toBeGreaterThan(1)
  })

  it('docx serialize/load preserves paragraph count', () => {
    let doc = createEmptyDocument('docx')
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'One',
    }).doc
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 3 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
      text: 'Two',
    }).doc
    const bytes = saveDocument(doc, 'docx')
    const loaded = loadDocument(bytes, 'docx')
    expect(loaded.sections[0]!.blocks.length).toBe(2)
  })
})
