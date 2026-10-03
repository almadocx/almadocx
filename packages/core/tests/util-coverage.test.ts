import { describe, expect, it } from 'vitest'
import { zipSync, strToU8 } from 'fflate'
import {
  AlmadocxError,
  assert,
  buildXml,
  createApproximateMeasurer,
  createEmptyDocument,
  detectFormat,
  fontCss,
  fontStackCss,
  inchesToTwips,
  isMacroPath,
  loadDocument,
  parseXml,
  pointsToTwips,
  pxToTwips,
  readZip,
  resolveFontFamily,
  saveDocument,
  twipsToPx,
  writeZip,
} from '../src/index.js'
import { unreachable } from '../src/util/assert.js'
import { hmToTwips, parseOdfLengthToTwips, twipsToHm } from '../src/util/units.js'
import { asArray, xmlAttr, xmlText } from '../src/io/xml.js'
import { setZipText, zipText } from '../src/io/zip.js'
import { PACKAGE_LIMITS } from '../src/io/limits.js'
import { mergeParagraphTabProps, defaultTabs, nextTabStopTwips } from '../src/layout/tabs.js'

describe('util units', () => {
  it('converts length units', () => {
    expect(inchesToTwips(1)).toBe(1440)
    expect(pointsToTwips(72)).toBe(1440)
    expect(twipsToPx(1440, 96)).toBe(96)
    expect(pxToTwips(96, 96)).toBe(1440)
    expect(hmToTwips(2540)).toBe(1440)
    expect(twipsToHm(1440)).toBe(2540)
  })

  it('parses ODF length strings', () => {
    expect(parseOdfLengthToTwips(undefined)).toBeUndefined()
    expect(parseOdfLengthToTwips('')).toBeUndefined()
    expect(parseOdfLengthToTwips('not-a-length')).toBeUndefined()
    expect(parseOdfLengthToTwips('2.54cm')).toBe(1440)
    expect(parseOdfLengthToTwips('25.4mm')).toBe(1440)
    expect(parseOdfLengthToTwips('1in')).toBe(1440)
    expect(parseOdfLengthToTwips('72pt')).toBe(1440)
    expect(parseOdfLengthToTwips('6pc')).toBe(1440)
    expect(parseOdfLengthToTwips('96px')).toBe(1440)
    expect(parseOdfLengthToTwips('12')).toBe(240) // default pt
  })
})

describe('assert helpers', () => {
  it('throws AlmadocxError on failed assert', () => {
    expect(() => assert(false, 'x', 'nope')).toThrow(AlmadocxError)
    assert(true, 'x', 'ok')
  })

  it('unreachable throws', () => {
    expect(() => unreachable('boom' as never)).toThrow(AlmadocxError)
  })
})

describe('fonts', () => {
  it('resolves substitutes and CSS stacks', () => {
    expect(resolveFontFamily(undefined)).toBe('Liberation Serif')
    expect(resolveFontFamily('"Calibri"')).toBe('Carlito')
    expect(resolveFontFamily('UnknownFace')).toBe('UnknownFace')
    expect(fontStackCss('Calibri')).toContain('Carlito')
    expect(fontStackCss('Calibri')).toContain('Liberation Sans')
    expect(fontStackCss('Liberation Serif')).toContain('DejaVu Serif')
    expect(fontCss('Arial', 12, true, true)).toMatch(/italic 700/)
    expect(fontCss('Arial', 12, false, false)).toMatch(/normal 400/)
  })
})

describe('measure', () => {
  it('measures wide and ascii characters', () => {
    const m = createApproximateMeasurer(0.5)
    const ascii = m.measure('ab', '16px serif')
    const wide = m.measure('你', '16px serif')
    expect(wide.width).toBeGreaterThan(ascii.width / 2)
    const fallback = m.measure('x', 'no-size')
    expect(fallback.width).toBeGreaterThan(0)
  })
})

describe('tabs helpers', () => {
  it('uses explicit tab stops and leaders', () => {
    const stop = nextTabStopTwips(0, [{ position: 1440, alignment: 'left', leader: 'dot' }], 5000)
    expect(stop.position).toBe(1440)
    expect(stop.leader).toBe('dot')
    const past = nextTabStopTwips(2000, [{ position: 1000, alignment: 'left' }], 3000)
    expect(past.position).toBeGreaterThan(2000)
    expect(defaultTabs()).toHaveLength(12)
    expect(mergeParagraphTabProps({ styleId: 'Normal' }, defaultTabs()).tabs).toHaveLength(12)
  })
})

describe('xml helpers', () => {
  it('builds and inspects xml nodes', () => {
    const parsed = parseXml('<root a="1"><child>hi</child></root>') as Record<string, unknown>
    expect(buildXml(parsed)).toContain('root')
    expect(asArray(undefined)).toEqual([])
    expect(asArray(1)).toEqual([1])
    expect(asArray([1, 2])).toEqual([1, 2])
    expect(xmlAttr(undefined, 'a')).toBeUndefined()
    expect(xmlAttr({ '@_a': 'x' }, 'a')).toBe('x')
    expect(xmlAttr({ '@_n': 3 }, 'n')).toBe('3')
    expect(xmlAttr({ '@_b': true }, 'b')).toBeUndefined()
    expect(xmlText(undefined)).toBe('')
    expect(xmlText(null)).toBe('')
    expect(xmlText('hi')).toBe('hi')
    expect(xmlText(7)).toBe('7')
    expect(xmlText({ '#text': 't' })).toBe('t')
    expect(xmlText({ '#text': 9 })).toBe('9')
    expect(xmlText({})).toBe('')
  })
})

describe('zip helpers', () => {
  it('writes and reads text entries; rejects bad paths', () => {
    const entries = new Map<string, Uint8Array>()
    setZipText(entries, 'a.txt', 'hello')
    const bytes = writeZip(entries)
    const read = readZip(bytes)
    expect(zipText(read, 'a.txt')).toBe('hello')
    expect(zipText(read, 'missing')).toBeUndefined()
    expect(isMacroPath('word/vbaProject.bin')).toBe(true)
    expect(isMacroPath('word/macrosheets/sheet1.xml')).toBe(true)
    expect(isMacroPath('word/embeddings/macro.bin')).toBe(true)
    expect(isMacroPath('Basic/Module.xml')).toBe(true)
    expect(isMacroPath('word/document.xml')).toBe(false)

    expect(() => readZip(zipSync({ 'C:/windows/x.txt': strToU8('x') }))).toThrow(AlmadocxError)
    expect(() => readZip(zipSync({ 'foo/./bar.txt': strToU8('x') }))).toThrow(AlmadocxError)
  })

  it('rejects oversized uncompressed entries via filter', () => {
    const huge = new Uint8Array(PACKAGE_LIMITS.maxUncompressedEntryBytes + 1)
    const bytes = zipSync({ 'big.bin': huge })
    expect(() => readZip(bytes)).toThrow(AlmadocxError)
  })

  it('rejects corrupt zip bytes', () => {
    expect(() => readZip(new Uint8Array([1, 2, 3, 4]))).toThrow(AlmadocxError)
  })
})

describe('format load/save branches', () => {
  it('detects formats and saves with hints', () => {
    const docx = saveDocument(createEmptyDocument('docx'), 'docx')
    const odt = saveDocument(createEmptyDocument('odt'), 'odt')
    expect(detectFormat(docx)).toBe('docx')
    expect(detectFormat(odt)).toBe('odt')
    expect(detectFormat(new Uint8Array([0, 1, 2]))).toBeUndefined()
    expect(loadDocument(docx).package.sourceFormat).toBe('docx')
    expect(loadDocument(odt).package.sourceFormat).toBe('odt')
    expect(() => loadDocument(new Uint8Array([0, 1, 2]))).toThrow(AlmadocxError)

    // Prefer ODT when both sniffers match (mimetype+content.xml and word/document.xml)
    const both = new Map<string, Uint8Array>()
    setZipText(both, 'mimetype', 'application/vnd.oasis.opendocument.text')
    setZipText(both, 'content.xml', `<?xml version="1.0"?><office:document-content/>`)
    setZipText(both, 'word/document.xml', `<?xml version="1.0"?><w:document/>`)
    expect(detectFormat(writeZip(both))).toBe('odt')

    const internal = createEmptyDocument('internal')
    expect(saveDocument(internal).byteLength).toBeGreaterThan(0)
    expect(() =>
      saveDocument(internal, 'unknown' as 'docx'),
    ).toThrow(AlmadocxError)
  })
})
