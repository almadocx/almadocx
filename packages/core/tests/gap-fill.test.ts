import { describe, expect, it, vi } from 'vitest'
import { strToU8 } from 'fflate'
import {
  AlmadocxError,
  applyOp,
  clampPosition,
  comparePositions,
  createEmptyDocument,
  deleteRangeInParagraph,
  documentCharCount,
  extractPlainRange,
  findAll,
  findNext,
  getParagraph,
  hitTestRun,
  insertTextInParagraph,
  layoutDocument,
  loadDocument,
  moveByWord,
  moveCaretByArrow,
  moveEnd,
  moveHome,
  moveLeft,
  moveRight,
  moveTableCellTab,
  moveTableCellVertical,
  moveVertical,
  nearestParagraphIndex,
  nextTabStopTwips,
  parseDocx,
  parseOdt,
  parseXml,
  patchLayoutCellParagraph,
  patchLayoutParagraph,
  renderLevelText,
  resolveListMarkers,
  resolveParagraphProps,
  resolveRunProps,
  saveDocument,
  serializeDocx,
  serializeOdt,
  setMarkInParagraph,
  type Document,
  type Table,
} from '../src/index.js'
import { parseNumberingXml } from '../src/formats/docx/numbering.js'
import { setZipText, writeZip } from '../src/io/zip.js'
import * as zipMod from '../src/io/zip.js'
import { parseOdfLengthToTwips } from '../src/util/units.js'

function para(text: string, id = 'p') {
  return {
    id,
    type: 'paragraph' as const,
    props: {},
    runs: [{ id: `r_${id}`, props: {}, content: { type: 'text' as const, text } }],
  }
}

function cell(text: string, extraProps: Record<string, unknown> = {}) {
  return {
    id: `c_${text}`,
    props: { ...extraProps },
    blocks: [para(text, `p_${text}`)],
  }
}

describe('bidi remaining branches', () => {
  it('handles RTL neutral ArrowLeft and empty/edge word moves', () => {
    expect(moveCaretByArrow('ש ם', 2, 'ArrowLeft', 'rtl')).toBe(1)
    expect(moveCaretByArrow('שלום', 2, 'ArrowLeft')).toBe(1)
    expect(moveCaretByArrow('', 0, 'ArrowLeft')).toBe(0)
    expect(moveByWord('   ', 3, -1)).toBe(0)
    expect(moveByWord('ab', 0, 1)).toBe(2)
    expect(moveByWord('a  b', 1, 1)).toBe(3)
  })
})

describe('navigation remaining branches', () => {
  it('covers empty-row tables, adjacent tables, tab at first cell, vertical edges', () => {
    const emptyRows: Document = createEmptyDocument()
    emptyRows.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [],
        rows: [{ id: 'r', props: {}, cells: [{ id: 'c', props: {}, blocks: [] }] }],
      },
    ]
    // firstCellInTable falls through to default path when no paragraphs
    expect(moveHome(emptyRows, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell).toEqual({
      row: 0,
      cell: 0,
      para: 0,
    })

    // single-cell table at doc end — exit right yields undefined path
    const solo = createEmptyDocument()
    solo.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [2000],
        rows: [{ id: 'r', props: {}, cells: [cell('X')] }],
      },
    ]
    expect(
      moveRight(solo, {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 1,
        cell: { row: 0, cell: 0, para: 0 },
      }).blockIndex,
    ).toBe(0)

    // Shift+Tab at first cell stays put
    expect(
      moveTableCellTab(
        solo,
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        -1,
      ).position.cell,
    ).toEqual({ row: 0, cell: 0, para: 0 })

    // Vertical at edge falls back to adjacentCell / stays
    expect(
      moveTableCellVertical(
        solo,
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        1,
      ).cell?.row,
    ).toBe(0)
    expect(
      moveTableCellVertical(
        solo,
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        -1,
      ).cell?.row,
    ).toBe(0)

    // moveVertical with bad section
    const noSect = createEmptyDocument()
    noSect.sections = []
    expect(moveVertical(noSect, { sectionIndex: 0, blockIndex: 0, offset: 0 }, 1).offset).toBe(0)

    // multi-row/col left/right within table (adjacentCell cell/row branches)
    const grid = createEmptyDocument()
    grid.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000, 1000],
        rows: [
          { id: 'r0', props: {}, cells: [cell('A'), cell('B')] },
          { id: 'r1', props: {}, cells: [cell('C'), cell('D')] },
        ],
      },
    ]
    expect(
      moveLeft(grid, {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 0,
        cell: { row: 0, cell: 1, para: 0 },
      }).cell?.cell,
    ).toBe(0)
    expect(
      moveLeft(grid, {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 0,
        cell: { row: 1, cell: 0, para: 0 },
      }).cell?.row,
    ).toBe(0)
    expect(moveEnd(grid, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell?.row).toBe(1)
  })
})

describe('findReplace remaining', () => {
  it('extracts across sections and wraps findNext', () => {
    const doc = createEmptyDocument()
    doc.sections.push({
      id: 's2',
      properties: structuredClone(doc.sections[0]!.properties),
      blocks: [para('beta', 'p2')],
    })
    doc.sections[0]!.blocks = [para('alpha', 'p1')]
    expect(
      extractPlainRange(doc, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 1 },
        focus: { sectionIndex: 1, blockIndex: 0, offset: 3 },
      }),
    ).toBe('lpha\nbet')

    // wrap findNext when after all matches
    const one = createEmptyDocument()
    one.sections[0]!.blocks = [para('foo bar foo', 'p')]
    const last = findNext(one, { query: 'foo' }, { sectionIndex: 0, blockIndex: 0, offset: 10 })
    expect(last?.range.anchor.offset).toBe(0)

    // wholeWord boundary chars
    expect(findAll(one, { query: 'foo', wholeWord: true })).toHaveLength(2)
    expect(findAll(createEmptyDocument(), { query: 'x' })).toEqual([])
  })
})

describe('util/io remaining', () => {
  it('hits unknown ODF unit, bad xml, zip directory entries, save format branches', async () => {
    expect(parseOdfLengthToTwips('1em')).toBeUndefined()

    const { XMLParser } = await import('fast-xml-parser')
    const spy = vi.spyOn(XMLParser.prototype, 'parse').mockImplementationOnce(() => {
      throw new Error('parse fail')
    })
    expect(() => parseXml('<root/>')).toThrow(AlmadocxError)
    spy.mockRestore()

    const entries = new Map<string, Uint8Array>()
    setZipText(entries, 'dir/', '')
    setZipText(entries, 'ok.txt', 'hi')
    const z = writeZip(entries)
    // directory-like keys may be skipped on read
    expect(zipMod.readZip(z).has('ok.txt')).toBe(true)

    const docxDoc = createEmptyDocument('docx')
    expect(saveDocument(docxDoc).byteLength).toBeGreaterThan(0)
    const odtDoc = createEmptyDocument('odt')
    expect(saveDocument(odtDoc).byteLength).toBeGreaterThan(0)

    expect(
      nextTabStopTwips(0, [{ position: 720, alignment: 'left' }], 5000).leader,
    ).toBe('none')
  })
})

describe('model remaining', () => {
  it('covers nearestParagraph, empty tables, hitTest, insert after atomic, marks on images', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = []
    expect(nearestParagraphIndex(doc, 0, 0)).toBe(0)
    expect(clampPosition(doc, { sectionIndex: 0, blockIndex: 0, offset: 5 })).toEqual({
      sectionIndex: 0,
      blockIndex: 0,
      offset: 0,
    })

    const emptyTable = createEmptyDocument()
    emptyTable.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [{ id: 'r', props: {}, cells: [] }],
      },
    ]
    expect(clampPosition(emptyTable, { sectionIndex: 0, blockIndex: 0, offset: 0 }).cell).toBeTruthy()

    const badParaCell = createEmptyDocument()
    badParaCell.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [
          {
            id: 'r',
            props: {},
            cells: [
              {
                id: 'c',
                props: {},
                // force non-paragraph via cast for clamp fallback
                blocks: [{ id: 'x', type: 'paragraph', props: {}, runs: [] } as never],
              },
            ],
          },
        ],
      },
    ]
    // empty runs — hitTest throws
    expect(() =>
      hitTestRun({ id: 'p', type: 'paragraph', props: {}, runs: [] }, 0),
    ).toThrow(AlmadocxError)

    const withImg = {
      id: 'p',
      type: 'paragraph' as const,
      props: {},
      runs: [
        {
          id: 'img',
          props: {},
          content: { type: 'image' as const, mediaId: 'm', widthTwips: 10, heightTwips: 10 },
        },
        { id: 't', props: {}, content: { type: 'text' as const, text: 'x' } },
      ],
    }
    expect(insertTextInParagraph(withImg, 1, '!', undefined).runs.length).toBeGreaterThan(1)
    const marked = setMarkInParagraph(withImg, 0, 1, { bold: true })
    expect(marked.runs[0]!.props.bold).toBe(true)
    const unmarked = setMarkInParagraph(withImg, 1, 2, { bold: true })
    expect(unmarked.runs[0]!.props.bold).toBeUndefined()

    expect(documentCharCount(createEmptyDocument())).toBe(0)
    expect(inlineDoc()).toBeTruthy()
  })
})

function inlineDoc() {
  const doc = createEmptyDocument()
  doc.sections[0]!.blocks = [para('hi')]
  return getParagraph(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 })
}

describe('styles cascade missing basedOn', () => {
  it('breaks when basedOn points nowhere', () => {
    const doc = createEmptyDocument()
    doc.styles.paragraphStyles['A'] = {
      id: 'A',
      name: 'A',
      basedOn: 'Missing',
      paragraph: {},
      character: { bold: true },
    }
    doc.styles.characterStyles['C'] = {
      id: 'C',
      name: 'C',
      basedOn: 'Gone',
      props: { italic: true },
    }
    const p = para('x')
    p.props.styleId = 'A'
    p.runs[0]!.styleId = 'C'
    doc.sections[0]!.blocks = [p]
    expect(resolveParagraphProps(doc, p).styleId).toBe('A')
    expect(resolveRunProps(doc, p, p.runs[0]!).italic).toBe(true)
  })
})

describe('numbering remaining branches', () => {
  it('skips broken instances and bullet without placeholders', () => {
    const doc = createEmptyDocument()
    doc.numbering.nums['orphan'] = { numId: 'orphan', abstractNumId: 'nope' }
    doc.numbering.abstractNums['bul'] = {
      id: 'bul',
      levels: [{ ilvl: 0, format: 'bullet', levelText: '◆', start: 1 }],
    }
    doc.numbering.nums['b1'] = { numId: 'b1', abstractNumId: 'bul' }
    doc.sections[0]!.blocks = [
      { ...para('a'), props: { numPr: { numId: 'orphan', ilvl: 0 } } },
      { ...para('b'), props: { numPr: { numId: 'b1', ilvl: 0 } } },
      { ...para('c'), props: { numPr: { numId: 'b1', ilvl: 99 } } },
    ]
    const markers = resolveListMarkers(doc, 0)
    expect(markers.get(1)?.text).toContain('◆')
  })
})

describe('docx parse remaining XML shapes', () => {
  it('parses hyperlinks, atLeast spacing, empty body, empty cells, alt headers', () => {
    const entries = new Map<string, Uint8Array>()
    setZipText(
      entries,
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    )
    setZipText(
      entries,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>` +
        `<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>` +
        `<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.png"/>` +
        `</Relationships>`,
    )
    setZipText(entries, 'word/media/pic.png', '\x89PNG')
    setZipText(
      entries,
      'word/header2.xml',
      `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:hdr>`,
    )
    setZipText(
      entries,
      'word/footer2.xml',
      `<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>F2</w:t></w:r></w:p></w:ftr>`,
    )
    setZipText(
      entries,
      'word/document.xml',
      `<?xml version="1.0"?>` +
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<w:body>` +
        `<w:p>` +
        `<w:pPr><w:spacing w:line="400" w:lineRule="atLeast"/><w:widowControl/></w:pPr>` +
        `<w:hyperlink r:id="x"><w:r><w:t>link</w:t></w:r></w:hyperlink>` +
        `<w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>c</w:t></w:r>` +
        `<w:r><w:rPr><w:color w:val="zzzzzz"/></w:rPr><w:t>z</w:t></w:r>` +
        `<w:r><w:br w:type="column"/></w:r>` +
        `<w:r><w:instrText> PAGE </w:instrText></w:r>` +
        `</w:p>` +
        `<w:tbl>` +
        `<w:tblPr><w:tblW w:w="10" w:type="pct"/><w:jc w:val="both"/>` +
        `<w:tblBorders><w:top w:val="nil"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>` +
        `<w:tr><w:tc>` +
        `<w:tcPr><w:tcW w:w="10" w:type="auto"/><w:vAlign w:val="top"/>` +
        `<w:tcBorders><w:top w:val="none"/></w:tcBorders><w:shd w:fill="auto"/></w:tcPr>` +
        `</w:tc></w:tr></w:tbl>` +
        `<w:sectPr/>` +
        `</w:body></w:document>`,
    )
    const parsed = parseDocx(writeZip(entries))
    expect(parsed.sections[0]!.header?.blocks[0]).toBeTruthy()
    expect(parsed.sections[0]!.footer?.blocks[0]).toBeTruthy()
    expect(parsed.sections[0]!.blocks.length).toBeGreaterThan(0)

    // empty body content → synthetic paragraph (sectPr only)
    const empty = new Map<string, Uint8Array>()
    setZipText(
      empty,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:body><w:sectPr/></w:body></w:document>`,
    )
    expect(parseDocx(writeZip(empty)).sections[0]!.blocks).toHaveLength(1)

    // upperRoman via numbering map
    const rom = parseNumberingXml(
      `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/>` +
        `<w:lvlText w:val="%1."/><w:start w:val="1"/><w:pPr><w:ind w:left="100" w:hanging="50"/></w:pPr>` +
        `<w:rPr><w:rFonts w:ascii="Times"/></w:rPr></w:lvl></w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="1"/>` +
        `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="2"/></w:lvlOverride></w:num>` +
        `<w:num w:numId="2"><w:abstractNumId/></w:num>` +
        `<w:abstractNum><w:lvl w:ilvl="0"/></w:abstractNum>` +
        `</w:numbering>`,
    )
    expect(rom.abstractNums['1']?.levels[0]?.format).toBe('upperRoman')
  })
})

describe('docx/odt serialize edge branches', () => {
  it('serializes table without gridCols and gridSpan>1; mocks zip failures', () => {
    const doc = createEmptyDocument('docx')
    doc.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [
          {
            id: 'r',
            props: {},
            cells: [
              {
                id: 'c',
                props: { gridSpan: 2 },
                blocks: [para('x')],
              },
            ],
          },
        ],
      },
    ]
    expect(serializeDocx(doc).byteLength).toBeGreaterThan(0)

    const odt = createEmptyDocument('odt')
    odt.sections[0]!.blocks = [para('o')]
    expect(serializeOdt(odt).byteLength).toBeGreaterThan(0)

    const spy = vi.spyOn(zipMod, 'writeZip').mockImplementation(() => {
      throw new Error('boom')
    })
    expect(() => serializeDocx(createEmptyDocument('docx'))).toThrow(AlmadocxError)
    expect(() => serializeOdt(createEmptyDocument('odt'))).toThrow(AlmadocxError)
    spy.mockRestore()
  })
})

describe('odt parse remaining', () => {
  it('parses spans, spaces, nested lists, empty tables, headers, empty body', () => {
    const entries = new Map<string, Uint8Array>()
    setZipText(entries, 'mimetype', 'application/vnd.oasis.opendocument.text')
    setZipText(
      entries,
      'styles.xml',
      `<?xml version="1.0"?>` +
        `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">` +
        `<office:master-styles><style:master-page style:name="Standard">` +
        `<style:header><text:p>Head</text:p></style:header>` +
        `<style:footer><text:p>Foot</text:p></style:footer>` +
        `</style:master-page></office:master-styles>` +
        `<office:styles>` +
        `<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>` +
        `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>` +
        `</office:styles></office:document-styles>`,
    )
    setZipText(
      entries,
      'content.xml',
      `<?xml version="1.0"?>` +
        `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
        `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
        `xmlns:xlink="http://www.w3.org/1999/xlink">` +
        `<office:automatic-styles>` +
        `<style:style style:name="AP" style:family="paragraph"><style:paragraph-properties fo:margin-left="1cm"/></style:style>` +
        `<style:style style:name="AT" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>` +
        `</office:automatic-styles>` +
        `<office:body><office:text>` +
        `<text:p text:style-name="P1">Hello<text:span text:style-name="T1">Bold</text:span>` +
        `<text:s text:c="3"/><text:tab/><text:line-break/>` +
        `<text:a xlink:href="http://x"><text:span>link</text:span></text:a></text:p>` +
        `<text:list text:style-name="Numbering"><text:list-item><text:p>One</text:p>` +
        `<text:list><text:list-item><text:p>Nested</text:p></text:list-item></text:list>` +
        `</text:list-item></text:list>` +
        `<table:table><table:table-row>` +
        `<table:table-cell table:number-columns-spanned="2"/>` +
        `<table:covered-table-cell/>` +
        `</table:table-row></table:table>` +
        `</office:text></office:body></office:document-content>`,
    )
    const doc = parseOdt(writeZip(entries))
    expect(doc.sections[0]!.header).toBeTruthy()
    expect(doc.sections[0]!.footer).toBeTruthy()
    expect(doc.sections[0]!.blocks.length).toBeGreaterThan(1)

    const empty = new Map<string, Uint8Array>()
    setZipText(empty, 'mimetype', 'application/vnd.oasis.opendocument.text')
    setZipText(
      empty,
      'content.xml',
      `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">` +
        `<office:body><office:text><text:sequence-decls/></office:text></office:body></office:document-content>`,
    )
    expect(parseOdt(writeZip(empty)).sections[0]!.blocks).toHaveLength(1)
  })
})

describe('layout remaining branches', () => {
  it('splits oversized paragraphs and tables without grid; patches edge cases', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.properties.pageSize = { width: 6120, height: 2000 }
    doc.sections[0]!.properties.margins = {
      top: 100,
      right: 100,
      bottom: 100,
      left: 100,
      header: 50,
      footer: 50,
    }
    // Very tall paragraph (many short lines via narrow width + long text) with widowControl off
    // so splitParagraphAcrossPages path can run when taller than page.
    doc.sections[0]!.blocks = [
      {
        id: 'tall',
        type: 'paragraph',
        props: { widowControl: false, keepLines: false },
        runs: [
          {
            id: 'r',
            props: { fontSizePt: 28 },
            content: { type: 'text', text: ('Word '.repeat(80) + '\n').repeat(40) },
          },
        ],
      },
    ]
    const tallLayout = layoutDocument(doc)
    expect(tallLayout.pages.length).toBeGreaterThan(0)

    // Table without gridCols + vAlign center/bottom with slack
    const tdoc = createEmptyDocument()
    tdoc.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: { borders: { top: '#ccc' } },
        rows: [
          {
            id: 'r0',
            props: { heightTwips: 2000 },
            cells: [
              {
                id: 'c0',
                props: {
                  vAlign: 'center',
                  vMerge: 'restart',
                  margin: { top: 20, bottom: 20, left: 10, right: 10 },
                },
                blocks: [para('hi')],
              },
              {
                id: 'c1',
                props: { vAlign: 'bottom', gridSpan: 1 },
                blocks: [para('lo'), para('lo2')],
              },
            ],
          },
          {
            id: 'r1',
            props: {},
            cells: [
              { id: 'c2', props: { vMerge: 'continue' }, blocks: [para('')] },
              { id: 'c3', props: {}, blocks: [para('x')] },
            ],
          },
        ],
      } satisfies Table,
    ]
    const tl = layoutDocument(tdoc)
    expect(tl.pages[0]!.blocks.some((b) => b.kind === 'table')).toBe(true)

    // patchLayoutParagraph early exits
    expect(patchLayoutParagraph(tdoc, tl, 0, 0).needsFullLayout).toBe(true)
    const pdoc = createEmptyDocument()
    pdoc.sections[0]!.blocks = [para('abc')]
    const pl = layoutDocument(pdoc)
    expect(patchLayoutParagraph(pdoc, pl, 0, 0).needsFullLayout).toBe(false)
    // force multi-occurrence by duplicating layout ref
    const splitPages = {
      pages: [
        { ...pl.pages[0]!, blocks: [...pl.pages[0]!.blocks, ...pl.pages[0]!.blocks] },
      ],
    }
    expect(patchLayoutParagraph(pdoc, splitPages as never, 0, 0).needsFullLayout).toBe(true)

    // patch cell on non-table / missing para
    expect(patchLayoutCellParagraph(pdoc, pl, 0, 0, { row: 0, cell: 0, para: 0 }).needsFullLayout).toBe(
      true,
    )
    expect(
      patchLayoutCellParagraph(tdoc, tl, 0, 0, { row: 9, cell: 0, para: 0 }).needsFullLayout,
    ).toBe(true)

    // right/center alignment + exact line spacing + list marker
    const fancy = createEmptyDocument()
    fancy.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: {
          alignment: 'right',
          lineSpacingRule: 'exact',
          lineSpacing: 480,
          numPr: { numId: '1', ilvl: 0 },
          indentLeft: 720,
        },
        runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'Right' } }],
      },
      {
        id: 'p2',
        type: 'paragraph',
        props: { alignment: 'center', keepNext: true, keepLines: true, pageBreakBefore: true },
        runs: [{ id: 'r2', props: {}, content: { type: 'text', text: 'Center next page' } }],
      },
    ]
    expect(layoutDocument(fancy).pages.length).toBeGreaterThan(0)
  })
})

describe('ops remaining branches', () => {
  it('covers more apply edge paths', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'abc',
    }).doc
    // merge at first paragraph throws
    expect(() =>
      applyOp(doc, {
        type: 'mergeParagraphs',
        position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      }),
    ).toThrow(AlmadocxError)

    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 1 },
    }).doc
    doc = applyOp(doc, {
      type: 'mergeParagraphs',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
    }).doc

    // setMark empty / inverted range handled
    doc = applyOp(doc, {
      type: 'setMark',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 2 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 1 },
      },
      mark: { bold: true },
    }).doc

    // insertColumn/deleteColumn restore paths already covered elsewhere — hit insertRow -1
    const tdoc = createEmptyDocument()
    tdoc.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000, 1000],
        rows: [
          { id: 'r0', props: {}, cells: [cell('A'), cell('B')] },
          { id: 'r1', props: {}, cells: [cell('C'), cell('D')] },
        ],
      } satisfies Table,
    ]
    doc = applyOp(tdoc, {
      type: 'insertRow',
      sectionIndex: 0,
      blockIndex: 0,
      afterRow: -1,
    }).doc
    expect((doc.sections[0]!.blocks[0] as Table).rows).toHaveLength(3)
    doc = applyOp(doc, {
      type: 'insertColumn',
      sectionIndex: 0,
      blockIndex: 0,
      afterCol: -1,
    }).doc
    expect((doc.sections[0]!.blocks[0] as Table).rows[0]!.cells.length).toBe(3)

    // ops on non-table block throw
    const plain = createEmptyDocument()
    expect(() =>
      applyOp(plain, { type: 'deleteRow', sectionIndex: 0, blockIndex: 0, row: 0 }),
    ).toThrow(/expected table/)
  })
})

describe('loadDocument formatHint', () => {
  it('loads with explicit hints', () => {
    const bytes = saveDocument(createEmptyDocument('docx'), 'docx')
    expect(loadDocument(bytes, 'docx').package.sourceFormat).toBe('docx')
    const odt = saveDocument(createEmptyDocument('odt'), 'odt')
    expect(loadDocument(odt, 'odt').package.sourceFormat).toBe('odt')
  })
})

describe('final gap sweep', () => {
  it('covers navigation row-wrap, vertical into table, multi-section ops', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [
      para('Before'),
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000, 1000],
        rows: [
          { id: 'r0', props: {}, cells: [cell('A'), cell('B')] },
          { id: 'r1', props: {}, cells: [cell('C'), cell('D')] },
        ],
      } satisfies Table,
      para('After'),
    ]
    // next-row wrap inside adjacentCell
    expect(
      moveRight(doc, {
        sectionIndex: 0,
        blockIndex: 1,
        offset: 1,
        cell: { row: 0, cell: 1, para: 0 },
      }).cell,
    ).toEqual({ row: 1, cell: 0, para: 0 })
    // vertical from last cell of last row falls back to adjacentCell (exit)
    expect(
      moveTableCellVertical(
        doc,
        { sectionIndex: 0, blockIndex: 1, offset: 0, cell: { row: 1, cell: 1, para: 0 } },
        1,
      ).blockIndex,
    ).toBe(2)
    // moveVertical into table from paragraph (preferOffset path)
    expect(
      moveVertical(doc, { sectionIndex: 0, blockIndex: 0, offset: 2 }, 1, 0).cell,
    ).toBeTruthy()
    // skip unknown block types then land on paragraph; fall off end
    const skipDoc = createEmptyDocument()
    ;(skipDoc.sections[0]!.blocks as unknown[]) = [
      para('a'),
      { id: 'u', type: 'unknown' },
      para('b'),
    ]
    expect(moveVertical(skipDoc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, 1).blockIndex).toBe(2)
    expect(moveVertical(skipDoc, { sectionIndex: 0, blockIndex: 2, offset: 0 }, 1).blockIndex).toBe(2)

    // multi-section apply ops (hits `: sec` map branches)
    const multi = createEmptyDocument()
    multi.sections.push({
      id: 's2',
      properties: structuredClone(multi.sections[0]!.properties),
      blocks: [para('sec2')],
    })
    multi.sections[0]!.blocks = [para('sec1a'), para('sec1b')]
    let d = applyOp(multi, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: '!',
    }).doc
    d = applyOp(d, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 0,
      numPr: { numId: '1', ilvl: 0 },
    }).doc
    d = applyOp(d, {
      type: 'setParagraphProps',
      sectionIndex: 0,
      blockIndex: 0,
      props: { alignment: 'left' },
    }).doc
    // set again to capture previous
    d = applyOp(d, {
      type: 'setParagraphProps',
      sectionIndex: 0,
      blockIndex: 0,
      props: { alignment: 'right' },
    }).doc
    d = applyOp(d, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 2 },
    }).doc
    d = applyOp(d, {
      type: 'mergeParagraphs',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
    }).doc
    expect(d.sections).toHaveLength(2)

    // insertTab after atomic: lone tab at offset=len → offsetInRun !== 0
    const tabOnly = createEmptyDocument()
    tabOnly.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: {},
        runs: [{ id: 't', props: {}, content: { type: 'tab' } }],
      },
    ]
    d = applyOp(tabOnly, {
      type: 'insertTab',
      position: { sectionIndex: 0, blockIndex: 0, offset: 1 },
    }).doc
    expect(d.sections[0]!.blocks[0]!.runs.length).toBeGreaterThan(1)

    // multi-section table ops hit replaceTable `: sec` branches
    const mt = createEmptyDocument()
    mt.sections.push({
      id: 's2',
      properties: structuredClone(mt.sections[0]!.properties),
      blocks: [para('other')],
    })
    mt.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000],
        rows: [
          { id: 'r0', props: {}, cells: [cell('A')] },
          { id: 'r1', props: {}, cells: [cell('B')] },
        ],
      } satisfies Table,
    ]
    d = applyOp(mt, {
      type: 'insertRow',
      sectionIndex: 0,
      blockIndex: 0,
      afterRow: 0,
    }).doc
    d = applyOp(d, {
      type: 'splitParagraph',
      position: {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 1,
        cell: { row: 0, cell: 0, para: 0 },
      },
    }).doc
    d = applyOp(d, {
      type: 'mergeParagraphs',
      position: {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 0,
        cell: { row: 0, cell: 0, para: 1 },
      },
    }).doc
    expect(d.sections).toHaveLength(2)

    // insertText after image (offsetInRun !== 0)
    const imgPara = {
      id: 'p',
      type: 'paragraph' as const,
      props: {},
      runs: [
        {
          id: 'img',
          props: {},
          content: { type: 'image' as const, mediaId: 'm', widthTwips: 10, heightTwips: 10 },
        },
      ],
    }
    expect(insertTextInParagraph(imgPara, 1, 'x').runs.length).toBe(2)

    // nearestParagraph skips unknown block types
    const weird = createEmptyDocument()
    weird.sections[0]!.blocks = [{ id: 'u', type: 'paragraph', props: {}, runs: [] } as never]
    // replace with invalid then para
    ;(weird.sections[0]!.blocks as unknown[]) = [
      { id: 'u', type: 'unknown' },
      para('ok'),
    ]
    expect(nearestParagraphIndex(weird, 0, 0)).toBe(1)
    expect(clampPosition(weird, { sectionIndex: 0, blockIndex: 0, offset: 0 }).blockIndex).toBe(1)

    // empty gridCols insertColumn uses 2400 default
    const g = createEmptyDocument()
    g.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [],
        rows: [{ id: 'r', props: {}, cells: [cell('A')] }],
      } satisfies Table,
    ]
    const g2 = applyOp(g, {
      type: 'insertColumn',
      sectionIndex: 0,
      blockIndex: 0,
      afterCol: 0,
    }).doc
    expect((g2.sections[0]!.blocks[0] as Table).gridCols?.[0]).toBe(2400)
  })

  it('covers docx/odt attribute variants and layout patch shifts', () => {
    // unprefixed numbering attrs + empty abstractNumId skip
    const n = parseNumberingXml(
      `<?xml version="1.0"?><numbering>` +
        `<abstractNum abstractNumId="a"><lvl ilvl="0">` +
        `<start val="2"/><numFmt val="decimal"/><lvlText val="%1."/>` +
        `<pPr><ind left="100" hanging="50"/></pPr>` +
        `<rPr><rFonts ascii="Arial"/></rPr>` +
        `</lvl></abstractNum>` +
        `<num numId="1"><abstractNumId val="a"/>` +
        `<lvlOverride ilvl="0"><startOverride val="9"/></lvlOverride></num>` +
        `</numbering>`,
    )
    expect(n.nums['1']?.startOverrides?.[0]).toBe(9)

    // rFonts hAnsi fallback + drawing path variants + webp + styles without defaults
    const entries = new Map<string, Uint8Array>()
    setZipText(
      entries,
      'word/styles.xml',
      `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>` +
        `<w:style w:type="character" w:styleId="Em"><w:name w:val="Em"/></w:style>` +
        `</w:styles>`,
    )
    setZipText(
      entries,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/media/x.webp"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/y.bmp"/>` +
        `</Relationships>`,
    )
    setZipText(entries, 'word/media/x.webp', 'WEBP')
    setZipText(entries, 'word/media/y.bmp', 'BM')
    setZipText(
      entries,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<w:body><w:p>` +
        `<w:r><w:rPr><w:rFonts hAnsi="Courier"/><w:b val="0"/><w:i val="false"/></w:rPr><w:t>x</w:t></w:r>` +
        `<w:r><w:drawing><wp:inline><wp:extent cx="635" cy="635"/>` +
        `<a:graphic><a:graphicData><a:blip r:embed="rId1"/></a:graphicData></a:graphic>` +
        `</wp:inline></w:drawing></w:r>` +
        `<w:r><w:drawing><wp:inline><a:blip r:embed="rId2"/></wp:inline></w:drawing></w:r>` +
        `<w:r><w:drawing><wp:inline><a:blip r:embed="nope"/></wp:inline></w:drawing></w:r>` +
        `</w:p><w:sectPr/></w:body></w:document>`,
    )
    const parsed = parseDocx(writeZip(entries))
    expect(Object.values(parsed.media).some((m) => m.contentType === 'image/webp')).toBe(true)
    expect(Object.values(parsed.media).some((m) => m.contentType === 'application/octet-stream')).toBe(
      true,
    )

    // serialize justify + invalid border colors
    const sdoc = createEmptyDocument('docx')
    sdoc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: { alignment: 'justify' },
        runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'j' } }],
      },
      {
        id: 't',
        type: 'table',
        props: { borders: { top: 'not-a-color', bottom: '#112233' } },
        rows: [
          {
            id: 'r',
            props: {},
            cells: [
              {
                id: 'c',
                props: { borders: { top: 'bad' } },
                blocks: [para('c')],
              },
            ],
          },
        ],
      },
    ]
    expect(serializeDocx(sdoc).byteLength).toBeGreaterThan(0)

    // odt nested inline s/tab/break + text node
    const odt = new Map<string, Uint8Array>()
    setZipText(odt, 'mimetype', 'application/vnd.oasis.opendocument.text')
    setZipText(
      odt,
      'content.xml',
      `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">` +
        `<office:body><office:text>` +
        `<text:p><text:span>A<text:s c="2"/><text:tab/><text:line-break/>B</text:span></text:p>` +
        `<text:p>plain-only</text:p>` +
        `<text:p><text:span/></text:p>` +
        `</office:text></office:body></office:document-content>`,
    )
    expect(parseOdt(writeZip(odt)).sections[0]!.blocks.length).toBeGreaterThan(1)

    // odt serialize image alt + preserved content.xml skip + existing styles.xml
    const odtDoc = createEmptyDocument('odt')
    odtDoc.package.preservedParts = [
      { path: 'content.xml', bytes: strToU8('old') },
      { path: 'styles.xml', bytes: strToU8('<styles/>') },
      { path: 'meta.xml', bytes: strToU8('<meta/>') },
    ]
    odtDoc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: { styleId: 'Standard' },
        runs: [
          {
            id: 'img',
            props: {},
            content: { type: 'image', mediaId: 'm', widthTwips: 10, heightTwips: 10, alt: 'pic' },
          },
        ],
      },
    ]
    expect(serializeOdt(odtDoc).byteLength).toBeGreaterThan(0)

    // serialize with left alignment (non-justify branch of jc mapping)
    const leftDoc = createEmptyDocument('docx')
    leftDoc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: { alignment: 'left' },
        runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'L' } }],
      },
    ]
    expect(serializeDocx(leftDoc).byteLength).toBeGreaterThan(0)

    // position: empty section blocks mid-clamp; hitTest past end with runs
    const emptySec = createEmptyDocument()
    emptySec.sections[0]!.blocks = []
    expect(clampPosition(emptySec, { sectionIndex: 0, blockIndex: 5, offset: 9 }).offset).toBe(0)
    const hitP = {
      id: 'p',
      type: 'paragraph' as const,
      props: {},
      runs: [{ id: 'r', props: {}, content: { type: 'text' as const, text: 'ab' } }],
    }
    expect(hitTestRun(hitP, 99).offsetInRun).toBe(2)

    // patchLayoutParagraph shift followers (para + table) and cell later paras
    const pdoc = createEmptyDocument()
    pdoc.sections[0]!.blocks = [
      para('short'),
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [4000],
        rows: [
          {
            id: 'r',
            props: {},
            cells: [
              {
                id: 'c',
                props: {},
                blocks: [para('c0', 'c0'), para('c1', 'c1')],
              },
            ],
          },
        ],
      } satisfies Table,
    ]
    let layout = layoutDocument(pdoc)
    // grow first paragraph so dy !== 0 and followers shift
    pdoc.sections[0]!.blocks[0] = para('short '.repeat(40))
    const patched = patchLayoutParagraph(pdoc, layout, 0, 0)
    expect(patched.layout.pages[0]).toBeTruthy()

    layout = layoutDocument(pdoc)
    ;(pdoc.sections[0]!.blocks[1] as Table).rows[0]!.cells[0]!.blocks[0] = para(
      'c0 '.repeat(30),
      'c0',
    )
    const cellPatched = patchLayoutCellParagraph(pdoc, layout, 0, 1, {
      row: 0,
      cell: 0,
      para: 0,
    })
    expect(cellPatched.layout.pages[0]).toBeTruthy()

    // patch two body paragraphs so follower paragraph shifts (1006-1007)
    const twoPara = createEmptyDocument()
    twoPara.sections[0]!.blocks = [para('aa'), para('bb')]
    let twoLayout = layoutDocument(twoPara)
    twoPara.sections[0]!.blocks[0] = para('aa '.repeat(50))
    expect(patchLayoutParagraph(twoPara, twoLayout, 0, 0).layout.pages[0]!.blocks.length).toBe(2)
    // patch second paragraph so earlier block hits `return b` (1012)
    twoLayout = layoutDocument(twoPara)
    twoPara.sections[0]!.blocks[1] = para('bb '.repeat(40))
    expect(patchLayoutParagraph(twoPara, twoLayout, 0, 1).needsFullLayout).toBeTypeOf('boolean')
    // dy≈0 path for later cell paragraph (1125)
    twoLayout = layoutDocument(pdoc)
    const same = patchLayoutCellParagraph(pdoc, twoLayout, 0, 1, { row: 0, cell: 0, para: 0 })
    expect(same.layout.pages[0]).toBeTruthy()
    // occurrences !== 1 for cell patch (1091-1092)
    const dup = layoutDocument(pdoc)
    if (dup.pages[0] && dup.pages[0].blocks[1]?.kind === 'table') {
      dup.pages.push({
        ...dup.pages[0],
        blocks: [dup.pages[0].blocks[1]!],
        paragraphs: [],
      })
    }
    expect(
      patchLayoutCellParagraph(pdoc, dup, 0, 1, { row: 0, cell: 0, para: 0 }).needsFullLayout,
    ).toBe(true)
    expect(
      patchLayoutCellParagraph(pdoc, { pages: [] }, 0, 1, { row: 0, cell: 0, para: 0 })
        .needsFullLayout,
    ).toBe(true)

    // position: only-unknown blocks; undefined hole; non-paragraph clamp
    const onlyUnknown = createEmptyDocument()
    ;(onlyUnknown.sections[0]!.blocks as unknown[]) = [
      { id: 'u0', type: 'unknown' },
      { id: 'u1', type: 'unknown' },
    ]
    expect(nearestParagraphIndex(onlyUnknown, 0, 1)).toBe(0)
    expect(clampPosition(onlyUnknown, { sectionIndex: 0, blockIndex: 0, offset: 3 }).offset).toBe(0)
    const hole = createEmptyDocument()
    hole.sections[0]!.blocks = [undefined as unknown as (typeof hole.sections)[0]['blocks'][0]]
    expect(clampPosition(hole, { sectionIndex: 0, blockIndex: 0, offset: 0 }).offset).toBe(0)

    // findReplace multi-section with hole + table skip
    const fr = createEmptyDocument()
    fr.sections = [
      fr.sections[0]!,
      undefined as unknown as (typeof fr.sections)[0],
      {
        id: 's3',
        properties: structuredClone(fr.sections[0]!.properties),
        blocks: [para('zz'), { id: 't', type: 'table', props: {}, rows: [] } as Table],
      },
    ]
    fr.sections[0]!.blocks = [para('aa')]
    expect(
      extractPlainRange(fr, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 2, blockIndex: 0, offset: 2 },
      }),
    ).toContain('aa')

    // isMacroPath .bin+macro branch
    expect(zipMod.isMacroPath('word/macros/foo.bin')).toBe(true)
  })
})

describe('branch coverage cleanup', () => {
  it('hits remaining parse/nav/layout/numbering branches', () => {
    // multi-section extract with block hole (findReplace:151)
    const fr = createEmptyDocument()
    fr.sections.push({
      id: 's2',
      properties: structuredClone(fr.sections[0]!.properties),
      blocks: [para('end')],
    })
    fr.sections[0]!.blocks = [
      para('start'),
      undefined as never,
      { id: 't', type: 'table', props: {}, rows: [] } as Table,
    ]
    expect(
      extractPlainRange(fr, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 1, blockIndex: 0, offset: 3 },
      }),
    ).toContain('start')

    // numbering deeper level gap → ?? 1
    const numDoc = createEmptyDocument()
    numDoc.numbering.abstractNums['g'] = {
      id: 'g',
      levels: [
        { ilvl: 0, format: 'decimal', levelText: '%1.', start: 1 },
        { ilvl: 2, format: 'decimal', levelText: '%3.', start: 7 },
      ],
    }
    numDoc.numbering.nums['g'] = { numId: 'g', abstractNumId: 'g' }
    numDoc.sections[0]!.blocks = [
      { ...para('a'), props: { numPr: { numId: 'g', ilvl: 0 } } },
      { ...para('b'), props: { numPr: { numId: 'g', ilvl: 2 } } },
      { ...para('c'), props: { numPr: { numId: 'g', ilvl: 0 } } },
    ]
    expect(resolveListMarkers(numDoc, 0).size).toBe(3)

    // firstCellInTable sparse row without cells
    const sparseNav = createEmptyDocument()
    sparseNav.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [undefined as never, { id: 'r', props: {}, cells: [cell('Z')] }],
      } satisfies Table,
    ]
    expect(moveHome(sparseNav, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell?.row).toBe(1)

    // docx: missing header xml, media endsWith match, vMerge continue/restart, margins, basedOn char
    const entries = new Map<string, Uint8Array>()
    setZipText(
      entries,
      'word/styles.xml',
      `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:style w:type="character" w:styleId="CBased"><w:name w:val="C"/>` +
        `<w:basedOn w:val="C1"/></w:style>` +
        `<w:style w:type="character" w:styleId="C1"><w:name w:val="C1"/></w:style>` +
        `</w:styles>`,
    )
    setZipText(
      entries,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdMiss" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerMissing.xml"/>` +
        `<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.png"/>` +
        `</Relationships>`,
    )
    // media stored under a path that only matches by filename endsWith
    setZipText(entries, 'word/media/other/pic.png', 'PNGDATA')
    // also seed media without path so path ?? '' runs during drawing resolve via preservedParts
    setZipText(
      entries,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
        `<w:body><w:p>` +
        `<w:r><w:drawing><wp:inline><a:blip r:embed="rIdImg"/></wp:inline></w:drawing></w:r>` +
        `<w:r><w:drawing><wp:inline><a:blip r:embed="nope"/></wp:inline></w:drawing></w:r>` +
        `</w:p>` +
        `<w:tbl><w:tr><w:tc><w:tcPr>` +
        `<w:vMerge w:val="continue"/>` +
        `<w:tcMar><w:top w:w="10"/><w:left/></w:tcMar>` +
        `</w:tcPr></w:tc>` +
        `<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc>` +
        `</w:tr></w:tbl>` +
        `<w:sectPr><w:pgMar w:top="100" w:left="100"/></w:sectPr>` +
        `</w:body></w:document>`,
    )
    const docx = parseDocx(writeZip(entries))
    expect(docx.styles.characterStyles['CBased']?.basedOn).toBe('C1')
    expect(docx.sections[0]!.properties.margins.top).toBe(100)

    // odt content without document-content wrapper
    const odt = new Map<string, Uint8Array>()
    setZipText(odt, 'mimetype', 'application/vnd.oasis.opendocument.text')
    setZipText(
      odt,
      'content.xml',
      `<?xml version="1.0"?><office:body xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">` +
        `<office:text><text:p>Bare</text:p></office:text></office:body>`,
    )
    expect(parseOdt(writeZip(odt)).sections[0]!.blocks[0]).toBeTruthy()

    // layout: auto lineSpacing default, tab without leader, break wrap, gridSpan overrun,
    // non-para cell block, marker shift, multi-page patch underflow
    const layDoc = createEmptyDocument()
    layDoc.sections[0]!.properties.pageSize = { width: 6120, height: 2800 }
    layDoc.sections[0]!.properties.margins = {
      top: 200,
      right: 200,
      bottom: 200,
      left: 200,
      header: 100,
      footer: 100,
    }
    layDoc.sections[0]!.blocks = [
      {
        id: 'list',
        type: 'paragraph',
        props: {
          numPr: { numId: '1', ilvl: 0 },
          lineSpacingRule: 'auto',
          // no lineSpacing → ?? 1.15
          // no spacingAfter → ?? 0
        },
        runs: [
          { id: 't', props: {}, content: { type: 'text', text: 'Item' } },
          { id: 'tab', props: {}, content: { type: 'tab' } },
          {
            id: 'long',
            props: {},
            content: { type: 'text', text: 'X'.repeat(200) },
          },
          {
            id: 'br',
            props: { strike: true, italic: true },
            content: { type: 'break', breakType: 'line' },
          },
          {
            id: 'br2',
            props: {},
            content: { type: 'break', breakType: 'line' },
          },
        ],
      },
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000],
        rows: [
          {
            id: 'r',
            props: {},
            cells: [
              {
                id: 'c',
                props: { gridSpan: 5 },
                blocks: [
                  para('c'),
                  { id: 'bad', type: 'unknown' } as never,
                ],
              },
            ],
          },
        ],
      } satisfies Table,
      para('after'),
    ]
    let layout = layoutDocument(layDoc)
    expect(layout.pages.length).toBeGreaterThan(0)
    // grow list then shrink on multi-page for underflow + marker shift of follower
    layDoc.sections[0]!.blocks[0] = {
      ...(layDoc.sections[0]!.blocks[0] as never),
      runs: [
        {
          id: 'r',
          props: {},
          content: { type: 'text', text: ('Word '.repeat(30) + '\n').repeat(25) },
        },
      ],
    } as never
    layout = layoutDocument(layDoc)
    expect(layout.pages.length).toBeGreaterThan(1)
    // shrink first para on page 0 while more pages exist
    layDoc.sections[0]!.blocks[0] = {
      id: 'list',
      type: 'paragraph',
      props: { numPr: { numId: '1', ilvl: 0 } },
      runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'short' } }],
    }
    const patched = patchLayoutParagraph(layDoc, layout, 0, 0)
    expect(patched.layout.pages.length).toBeGreaterThan(1)

    // force splitParagraphAcrossPages with prior content on page (line 239)
    const splitDoc = createEmptyDocument()
    splitDoc.sections[0]!.properties.pageSize = { width: 6120, height: 2000 }
    splitDoc.sections[0]!.properties.margins = {
      top: 100,
      right: 100,
      bottom: 100,
      left: 100,
      header: 50,
      footer: 50,
    }
    splitDoc.sections[0]!.blocks = [
      para('before'),
      {
        id: 'tall',
        type: 'paragraph',
        props: { widowControl: false, keepLines: false },
        runs: [
          {
            id: 'r',
            props: { fontSizePt: 28 },
            content: { type: 'text', text: ('Word '.repeat(20) + '\n').repeat(30) },
          },
        ],
      },
    ]
    expect(layoutDocument(splitDoc).pages.length).toBeGreaterThan(1)

    // sparse undefined cell in firstCellInTable
    const holeCell = createEmptyDocument()
    holeCell.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [{ id: 'r', props: {}, cells: [undefined as never, cell('ok')] }],
      } satisfies Table,
    ]
    expect(moveHome(holeCell, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell?.cell).toBe(1)

    // ghost image (rels target with no media bytes) + empty-attr tcMar
    const ghost = new Map<string, Uint8Array>()
    setZipText(
      ghost,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdG" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/ghost.png"/>` +
        `</Relationships>`,
    )
    setZipText(
      ghost,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<w:body><w:p><w:r><w:drawing><a:blip r:embed="rIdG"/></w:drawing></w:r></w:p>` +
        `<w:tbl><w:tr><w:tc><w:tcPr><w:tcMar><w:top/><w:left/></w:tcMar></w:tcPr>` +
        `<w:p/></w:tc></w:tr></w:tbl>` +
        `<w:sectPr/></w:body></w:document>`,
    )
    const ghostDoc = parseDocx(writeZip(ghost))
    expect(Object.values(ghostDoc.media).some((m) => m.bytes.byteLength === 0)).toBe(true)

    // marker shift: list follower moves when prior para shrinks; multi-page underflow
    const markDoc = createEmptyDocument()
    markDoc.sections[0]!.blocks = [
      para('AAAAAAAAAA'),
      {
        id: 'li',
        type: 'paragraph',
        props: { numPr: { numId: '1', ilvl: 0 } },
        runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'listed' } }],
      },
      {
        id: 'pg',
        type: 'paragraph',
        props: { pageBreakBefore: true },
        runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'page2' } }],
      },
    ]
    let markLay = layoutDocument(markDoc)
    expect(markLay.pages.length).toBeGreaterThan(1)
    // grow then shrink first para so dy < 0 with followers (marker) on same page
    markDoc.sections[0]!.blocks[0] = para('AA '.repeat(80))
    markLay = layoutDocument(markDoc)
    markDoc.sections[0]!.blocks[0] = para('x')
    const markPatched = patchLayoutParagraph(markDoc, markLay, 0, 0)
    expect(markPatched.layout.pages.length).toBeGreaterThan(0)

    // tab with explicit leader vs default none
    const tabDoc = createEmptyDocument()
    tabDoc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: {
          tabs: [
            { position: 1440, alignment: 'left', leader: 'dot' },
            { position: 2880, alignment: 'left' },
          ],
          spacingAfter: 120,
          lineSpacingRule: 'auto',
          lineSpacing: 1.5,
        },
        runs: [
          { id: 'a', props: {}, content: { type: 'text', text: 'A' } },
          { id: 't1', props: {}, content: { type: 'tab' } },
          { id: 'b', props: {}, content: { type: 'text', text: 'B' } },
          { id: 't2', props: {}, content: { type: 'tab' } },
          { id: 'c', props: {}, content: { type: 'text', text: 'C' } },
          // fill line then break to hit flush before break
          { id: 'fill', props: {}, content: { type: 'text', text: 'Z'.repeat(120) } },
          {
            id: 'br',
            props: { italic: true, strike: true },
            content: { type: 'break', breakType: 'line' },
          },
        ],
      },
    ]
    expect(layoutDocument(tabDoc).pages[0]!.paragraphs.length).toBeGreaterThan(0)

    // Clear paragraph defaults so spacingAfter/lineSpacing ?? arms run; force break wrap flush
    const defDoc = createEmptyDocument()
    defDoc.styles.docDefaults.paragraph = {}
    defDoc.sections[0]!.properties.pageSize = { width: 2000, height: 5000 }
    defDoc.sections[0]!.properties.margins = {
      top: 100,
      right: 100,
      bottom: 100,
      left: 100,
      header: 50,
      footer: 50,
    }
    defDoc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: { lineSpacingRule: 'auto' },
        runs: [
          { id: 't', props: {}, content: { type: 'text', text: 'WWWWWWWWWWWWWWWWWWWW' } },
          {
            id: 'br',
            props: {},
            content: { type: 'break', breakType: 'line' },
          },
          { id: 't2', props: {}, content: { type: 'text', text: 'more' } },
        ],
      },
      {
        id: 'p2',
        type: 'paragraph',
        props: { spacingAfter: 50, lineSpacingRule: 'auto', lineSpacing: 2 },
        runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'spaced' } }],
      },
    ]
    expect(layoutDocument(defDoc).pages[0]!.paragraphs.length).toBe(2)
  })
})

describe('branch coverage final push', () => {
  it('covers zip macro module.xml, xmlAttrAltOr, numbering root fallback', () => {
    expect(zipMod.isMacroPath('Scripts/xbasic/module.xml')).toBe(true)
    expect(zipMod.isMacroPath('word/module.xml')).toBe(false)

    // numbering without <numbering> wrapper + missing level attrs (defaults)
    const n = parseNumberingXml(
      `<?xml version="1.0"?><abstractNum abstractNumId="z"><lvl>` +
        `<numFmt val="decimal"/><lvlText/><pPr><ind/></pPr></lvl></abstractNum>` +
        `<num numId="9"><abstractNumId val="z"/>` +
        `<lvlOverride><startOverride val="3"/></lvlOverride></num>`,
    )
    expect(n.abstractNums['z']?.levels[0]?.ilvl).toBe(0)
    expect(n.abstractNums['z']?.levels[0]?.levelText).toBe('%1.')
    expect(n.abstractNums['z']?.levels[0]?.indentLeft).toBe(720)
    expect(n.nums['9']?.startOverrides?.[0]).toBe(3)

    // lvl without ind → ternary false arm
    const n2 = parseNumberingXml(
      `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:abstractNum w:abstractNumId="q"><w:lvl w:ilvl="0">` +
        `<w:numFmt w:val="bullet"/><w:lvlText w:val="*"/></w:lvl></w:abstractNum>` +
        `</w:numbering>`,
    )
    expect(n2.abstractNums['q']?.levels[0]?.indentLeft).toBe(720)
  })

  it('covers findReplace holes, table extract, non-para same-block', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [
      para('findme'),
      undefined as unknown as (typeof doc.sections)[0]['blocks'][0],
      para('other'),
    ]
    expect(findAll(doc, { query: 'find' })).toHaveLength(1)
    expect(
      extractPlainRange(doc, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 2, offset: 2 },
      }),
    ).toContain('findme')

    // table with sparse holes + reverse selection
    const tdoc = createEmptyDocument()
    const table: Table = {
      id: 't',
      type: 'table',
      props: {},
      rows: [
        {
          id: 'r0',
          props: {},
          cells: [
            {
              id: 'c0',
              props: {},
              blocks: [para('A0', 'a0'), undefined as never, para('A1', 'a1')],
            },
            undefined as never,
            { id: 'c1', props: {}, blocks: [para('B', 'b')] },
          ],
        },
        undefined as never,
        {
          id: 'r1',
          props: {},
          cells: [{ id: 'c2', props: {}, blocks: [para('C', 'c')] }],
        },
      ],
    }
    tdoc.sections[0]!.blocks = [table]
    expect(
      extractPlainRange(tdoc, {
        anchor: {
          sectionIndex: 0,
          blockIndex: 0,
          offset: 1,
          cell: { row: 0, cell: 0, para: 0 },
        },
        focus: {
          sectionIndex: 0,
          blockIndex: 0,
          offset: 1,
          cell: { row: 2, cell: 0, para: 0 },
        },
      }),
    ).toBeTruthy()
    // reverse endpoints
    expect(
      extractPlainRange(tdoc, {
        focus: {
          sectionIndex: 0,
          blockIndex: 0,
          offset: 0,
          cell: { row: 0, cell: 0, para: 0 },
        },
        anchor: {
          sectionIndex: 0,
          blockIndex: 0,
          offset: 1,
          cell: { row: 0, cell: 2, para: 0 },
        },
      }),
    ).toContain('B')

    // same-block non-paragraph / non-table → ''
    const weird = createEmptyDocument()
    ;(weird.sections[0]!.blocks as unknown[]) = [{ id: 'u', type: 'unknown' }]
    expect(
      extractPlainRange(weird, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 1 },
      }),
    ).toBe('')

    // missing block hole same-block
    const hole = createEmptyDocument()
    hole.sections[0]!.blocks = [undefined as never]
    expect(
      extractPlainRange(hole, {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        focus: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      }),
    ).toBe('')
  })

  it('covers navigation sparse rows, empty sections, non-table tab/vertical', () => {
    const sparse = createEmptyDocument()
    sparse.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000, 1000],
        rows: [
          undefined as never,
          {
            id: 'r1',
            props: {},
            cells: [cell('X'), cell('Y')],
          },
        ],
      } satisfies Table,
    ]
    expect(moveHome(sparse, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).cell?.row).toBe(1)

    // moveEnd with empty sections
    const empty = createEmptyDocument()
    empty.sections = []
    expect(moveEnd(empty, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).offset).toBe(0)

    // moveEnd skips trailing non-paragraph
    const skip = createEmptyDocument()
    ;(skip.sections[0]!.blocks as unknown[]) = [para('a'), { id: 'u', type: 'unknown' }]
    expect(moveEnd(skip, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true).blockIndex).toBe(0)

    // tab/vertical on non-table
    const plain = createEmptyDocument()
    plain.sections[0]!.blocks = [para('p')]
    expect(
      moveTableCellTab(plain, { sectionIndex: 0, blockIndex: 0, offset: 0 }, 1).position.offset,
    ).toBe(0)
    expect(
      moveTableCellVertical(plain, { sectionIndex: 0, blockIndex: 0, offset: 0 }, 1).offset,
    ).toBe(0)

    // tab next-cell with sparse row length ?? 0
    expect(
      moveTableCellTab(
        sparse,
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 1, cell: 0, para: 0 } },
        1,
      ).position.cell?.cell,
    ).toBe(1)
  })

  it('covers model text styleId, marks on breaks, position compare/holes', () => {
    const styled = {
      id: 'p',
      type: 'paragraph' as const,
      props: {},
      runs: [
        {
          id: 'r',
          styleId: 'Em',
          props: { bold: true },
          content: { type: 'text' as const, text: 'ab' },
        },
        { id: 't', props: {}, content: { type: 'tab' as const } },
        {
          id: 'br',
          props: { italic: true },
          content: { type: 'break' as const, breakType: 'line' as const },
        },
      ],
    }
    expect(insertTextInParagraph(styled, 1, 'X').runs.some((r) => r.styleId === 'Em')).toBe(true)
    expect(deleteRangeInParagraph(styled, 0, 1).runs[0]?.styleId).toBe('Em')
    // mark covering tab (inside) and partial (outside) on break
    const marked = setMarkInParagraph(styled, 2, 3, { underline: true })
    expect(marked.runs.some((r) => r.content.type === 'tab' && r.props.underline)).toBe(true)
    const markedOut = setMarkInParagraph(styled, 0, 1, { strike: true })
    expect(markedOut.runs.some((r) => r.content.type === 'break')).toBe(true)

    // comparePositions across sections and cell vs no-cell
    expect(
      comparePositions(
        { sectionIndex: 0, blockIndex: 0, offset: 0 },
        { sectionIndex: 1, blockIndex: 0, offset: 0 },
      ),
    ).toBeLessThan(0)
    expect(
      comparePositions(
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        { sectionIndex: 0, blockIndex: 0, offset: 0 },
      ),
    ).toBeGreaterThan(0)

    // nearestParagraph search backward to table; sparse firstCellPath
    const back = createEmptyDocument()
    ;(back.sections[0]!.blocks as unknown[]) = [
      { id: 't', type: 'table', props: {}, rows: [{ id: 'r', props: {}, cells: [] }] },
      { id: 'u', type: 'unknown' },
    ]
    expect(nearestParagraphIndex(back, 0, 1)).toBe(0)
    const sparseT = createEmptyDocument()
    sparseT.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [
          undefined as never,
          { id: 'r', props: {}, cells: [undefined as never, { id: 'c', props: {}, blocks: [] }] },
        ],
      } satisfies Table,
    ]
    expect(clampPosition(sparseT, { sectionIndex: 0, blockIndex: 0, offset: 0 }).cell).toBeTruthy()

    // hitTestRun with hole in runs
    const holeRuns = {
      id: 'p',
      type: 'paragraph' as const,
      props: {},
      runs: [undefined as never, { id: 'r', props: {}, content: { type: 'text' as const, text: 'z' } }],
    }
    expect(hitTestRun(holeRuns, 0).run.content).toMatchObject({ text: 'z' })
  })

  it('covers model numbering edge markers and apply branches', () => {
    const doc = createEmptyDocument()
    doc.numbering.abstractNums['empty'] = { id: 'empty', levels: [] }
    doc.numbering.nums['e'] = { numId: 'e', abstractNumId: 'empty' }
    doc.numbering.abstractNums['bul'] = {
      id: 'bul',
      levels: [
        { ilvl: 0, format: 'bullet', levelText: '', start: 1 },
        { ilvl: 1, format: 'decimal', levelText: '%1.', start: 5 },
      ],
    }
    doc.numbering.nums['b'] = { numId: 'b', abstractNumId: 'bul' }
    // missing instance + empty levels + bullet empty levelText + deeper reset
    doc.sections[0]!.blocks = [
      { ...para('x'), props: { numPr: { numId: 'missing', ilvl: 0 } } },
      { ...para('e'), props: { numPr: { numId: 'e', ilvl: 0 } } },
      { ...para('b0'), props: { numPr: { numId: 'b', ilvl: 0 } } },
      { ...para('b1'), props: { numPr: { numId: 'b', ilvl: 1 } } },
      { ...para('b0b'), props: { numPr: { numId: 'b', ilvl: 0 } } },
    ]
    const markers = resolveListMarkers(doc, 0)
    expect(markers.get(2)?.text).toBe('•')
    expect(markers.has(1)).toBe(false)

    // renderLevelText counters hole / missing level start
    expect(renderLevelText('%1.', [], [{ ilvl: 0, format: 'decimal', levelText: '%1.', start: 9 }])).toBe(
      '9.',
    )
    expect(renderLevelText('%2.', [1], [])).toBe('1.')

    // apply: multi-para cell update (map : p arm), empty runs delete props, widthTwips row
    let d = createEmptyDocument()
    d.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [1000, 1000],
        rows: [
          {
            id: 'r',
            props: {},
            cells: [
              {
                id: 'c',
                props: { widthTwips: 900 },
                blocks: [para('one', 'p0'), para('two', 'p1')],
              },
              cell('B'),
            ],
          },
        ],
      } satisfies Table,
      para('sibling'),
    ]
    d = applyOp(d, {
      type: 'insertText',
      position: {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 0,
        cell: { row: 0, cell: 0, para: 1 },
      },
      text: '!',
    }).doc
    expect(getParagraph(d, { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 1 } }).runs[0])
      .toBeTruthy()

    d = applyOp(d, {
      type: 'insertRow',
      sectionIndex: 0,
      blockIndex: 0,
      afterRow: 0,
    }).doc
    expect((d.sections[0]!.blocks[0] as Table).rows[1]!.cells[0]!.props.widthTwips).toBe(900)

    // delete with empty runs → ?? {}
    const emptyRuns = createEmptyDocument()
    emptyRuns.sections[0]!.blocks = [
      { id: 'p', type: 'paragraph', props: {}, runs: [] },
    ]
    expect(() =>
      applyOp(emptyRuns, {
        type: 'deleteRange',
        range: {
          anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
          focus: { sectionIndex: 0, blockIndex: 0, offset: 0 },
        },
      }),
    ).not.toThrow()

    // multi-section cross-paragraph delete
    const multi = createEmptyDocument()
    multi.sections.push({
      id: 's2',
      properties: structuredClone(multi.sections[0]!.properties),
      blocks: [para('keep')],
    })
    multi.sections[0]!.blocks = [para('AAAA'), para('BBBB')]
    const del = applyOp(multi, {
      type: 'deleteRange',
      range: {
        anchor: { sectionIndex: 0, blockIndex: 0, offset: 1 },
        focus: { sectionIndex: 0, blockIndex: 1, offset: 2 },
      },
    })
    expect(del.doc.sections).toHaveLength(2)

    // split/merge in table with sibling block (map : b arm)
    let mt = createEmptyDocument()
    mt.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [{ id: 'r', props: {}, cells: [{ id: 'c', props: {}, blocks: [para('xy', 'p')] }] }],
      } satisfies Table,
      para('after'),
    ]
    mt = applyOp(mt, {
      type: 'splitParagraph',
      position: {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 1,
        cell: { row: 0, cell: 0, para: 0 },
      },
    }).doc
    mt = applyOp(mt, {
      type: 'mergeParagraphs',
      position: {
        sectionIndex: 0,
        blockIndex: 0,
        offset: 0,
        cell: { row: 0, cell: 0, para: 1 },
      },
    }).doc
    expect(mt.sections[0]!.blocks).toHaveLength(2)

    // deleteColumn without gridCols (deletedGridCol undefined spreads)
    let nog = createEmptyDocument()
    nog.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [
          { id: 'r', props: {}, cells: [cell('A'), cell('B')] },
        ],
      } satisfies Table,
    ]
    nog = applyOp(nog, {
      type: 'deleteColumn',
      sectionIndex: 0,
      blockIndex: 0,
      col: 1,
    }).doc
    expect((nog.sections[0]!.blocks[0] as Table).rows[0]!.cells).toHaveLength(1)

    // deleteColumn on empty-rows hits ?? 0 then assert
    const emptyRows = createEmptyDocument()
    emptyRows.sections[0]!.blocks = [
      { id: 't', type: 'table', props: {}, rows: [] } satisfies Table,
    ]
    expect(() =>
      applyOp(emptyRows, { type: 'deleteColumn', sectionIndex: 0, blockIndex: 0, col: 0 }),
    ).toThrow()
  })

  it('covers docx parse remaining XML edge branches', () => {
    const entries = new Map<string, Uint8Array>()
    // Relationships without Relationships wrapper; styles without styles wrapper
    setZipText(
      entries,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/>`,
    )
    setZipText(entries, 'word/media/a.png', 'PNG')
    setZipText(
      entries,
      'word/styles.xml',
      `<?xml version="1.0"?><style type="paragraph" styleId=""><w:name w:val="Bad"/></style>` +
        `<style w:type="character" w:styleId="C1"><w:name w:val="C"/></style>`,
    )
    setZipText(
      entries,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<w:body>` +
        `<w:p><w:pPr><w:numPr><w:ilvl/><w:numId w:val="1"/></w:numPr>` +
        `<w:widowControl w:val="false"/></w:pPr>` +
        `<w:r><w:instrText>PAGE\\* MERGEFORMAT</w:instrText></w:r>` +
        `<w:r><w:drawing><a:blip embed="rId1"/></w:drawing></w:r>` +
        `<w:r><w:drawing><a:graphic/></w:drawing></w:r>` +
        `</w:p>` +
        `<w:p/>` +
        `<w:tbl>` +
        `<w:tr><w:tc><w:tcPr>` +
        `<w:tcBorders><w:top w:val="single"/></w:tcBorders>` +
        `<w:vMerge/><w:tcMar/></w:tcPr></w:tc></w:tr>` +
        `</w:tbl>` +
        `<w:sectPr><w:pgMar/></w:sectPr>` +
        `</w:body></w:document>`,
    )
    const parsed = parseDocx(writeZip(entries))
    expect(parsed.sections[0]!.blocks.length).toBeGreaterThan(0)

    // header/footer xml undefined path already; root-as-hf via unknown root element
    const hf = new Map<string, Uint8Array>()
    setZipText(
      hf,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="word/headerX.xml"/>` +
        `</Relationships>`,
    )
    setZipText(
      hf,
      'word/headerX.xml',
      `<?xml version="1.0"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>HX</w:t></w:r></w:p>`,
    )
    setZipText(
      hf,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    )
    expect(parseDocx(writeZip(hf)).sections[0]!.header?.blocks.length).toBeGreaterThan(0)
  })

  it('covers odt parse fallbacks and serialize image without alt', () => {
    const entries = new Map<string, Uint8Array>()
    setZipText(entries, 'mimetype', 'application/vnd.oasis.opendocument.text')
    // styles without document-styles wrapper
    setZipText(
      entries,
      'styles.xml',
      `<?xml version="1.0"?><style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
        `style:name="P" style:family="paragraph"/>` +
        `<style:style style:family="text"/>`,
    )
    // content without document-content wrapper — still need body/text under root for assert
    setZipText(
      entries,
      'content.xml',
      `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">` +
        `<office:automatic-styles>` +
        `<style:style style:family="paragraph"/>` +
        `<style:style style:name="T" style:family="text"/>` +
        `</office:automatic-styles>` +
        `<office:body><office:text>` +
        `<text:p><text:span text:style-name="Missing">x</text:span><text:s/></text:p>` +
        `<table:table><table:table-row><table:table-header-rows/>` +
        `<table:table-cell/><table:covered-table-cell/></table:table-row></table:table>` +
        `</office:text></office:body></office:document-content>`,
    )
    expect(parseOdt(writeZip(entries)).sections[0]!.blocks.length).toBeGreaterThan(0)

    const odtDoc = createEmptyDocument('odt')
    odtDoc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: {},
        runs: [
          {
            id: 'img',
            props: {},
            content: { type: 'image', mediaId: 'm', widthTwips: 10, heightTwips: 10 },
          },
        ],
      },
    ]
    const bytes = serializeOdt(odtDoc)
    const z = zipMod.readZip(bytes)
    expect(zipMod.zipText(z, 'content.xml')).toContain('[image]')
  })

  it('covers layout remaining: titlePage fallbacks, marks, patch multi-page', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.properties.titlePage = true
    // headerFirst missing → fall back to header; same for footer
    doc.sections[0]!.header = {
      blocks: [para('H')],
    }
    doc.sections[0]!.footer = {
      blocks: [para('F')],
    }
    doc.styles.docDefaults.character = {}
    doc.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: {
          spacingAfter: 200,
          lineSpacingRule: 'auto',
          lineSpacing: 2,
          indentFirstLine: 240,
          numPr: { numId: '1', ilvl: 0 },
        },
        runs: [
          {
            id: 'r',
            props: { italic: true, underline: true, strike: true, fontSizePt: 14, color: '#123456' },
            content: { type: 'text', text: 'HelloWorld'.repeat(40) },
          },
          { id: 'tab', props: { italic: true }, content: { type: 'tab' } },
          {
            id: 'br',
            props: { italic: true, underline: true, strike: true },
            content: { type: 'break', breakType: 'line' },
          },
          {
            id: 'img',
            props: {},
            content: { type: 'image', mediaId: 'm', widthTwips: 2000, heightTwips: 2000 },
          },
          {
            id: 'img2',
            props: {},
            content: {
              type: 'image',
              mediaId: 'm',
              widthTwips: 2000,
              heightTwips: 2000,
              alt: 'x',
            },
          },
        ],
      },
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [0, 0],
        rows: [
          {
            id: 'r0',
            props: { heightTwips: 400 },
            cells: [
              {
                id: 'c0',
                props: { vMerge: 'restart', vAlign: 'center', gridSpan: 2 },
                blocks: [para('hi'), { id: 'x', type: 'paragraph', props: {}, runs: [] } as never],
              },
            ],
          },
          {
            id: 'r1',
            props: {},
            cells: [
              { id: 'c1', props: { vMerge: 'continue', gridSpan: 2 }, blocks: [para('')] },
            ],
          },
          {
            id: 'r2',
            props: {},
            // break merge chain
            cells: [{ id: 'c2', props: {}, blocks: [para('z')] }],
          },
        ],
      } satisfies Table,
    ]
    // force marker path with defaults
    doc.numbering = createEmptyDocument().numbering
    const layout = layoutDocument(doc)
    expect(layout.pages.length).toBeGreaterThan(0)

    // patch with multi-page + shrink for underflow
    const tall = createEmptyDocument()
    tall.sections[0]!.properties.pageSize = { width: 6120, height: 2500 }
    tall.sections[0]!.properties.margins = {
      top: 200,
      right: 200,
      bottom: 200,
      left: 200,
      header: 100,
      footer: 100,
    }
    tall.sections[0]!.blocks = [
      {
        id: 'p',
        type: 'paragraph',
        props: { widowControl: false },
        runs: [
          {
            id: 'r',
            props: { fontSizePt: 20 },
            content: { type: 'text', text: ('Line\n').repeat(80) },
          },
        ],
      },
    ]
    let lay = layoutDocument(tall)
    expect(lay.pages.length).toBeGreaterThan(1)
    tall.sections[0]!.blocks[0] = para('short')
    const patched = patchLayoutParagraph(tall, lay, 0, 0)
    expect(patched.needsFullLayout).toBeTypeOf('boolean')

    // empty header/footer band pageNumberText fallback — header with no paras
    const emptyHf = createEmptyDocument()
    emptyHf.sections[0]!.header = { blocks: [] }
    emptyHf.sections[0]!.blocks = [para('x')]
    expect(layoutDocument(emptyHf).pages[0]!.header).toBeTruthy()

    // vAlign with no slack (tall content in short row skipped)
    const tight = createEmptyDocument()
    tight.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        rows: [
          {
            id: 'r',
            props: { heightTwips: 1 },
            cells: [
              {
                id: 'c',
                props: { vAlign: 'bottom' },
                blocks: [
                  {
                    id: 'p',
                    type: 'paragraph',
                    props: {},
                    runs: [
                      {
                        id: 'r',
                        props: { fontSizePt: 48 },
                        content: { type: 'text', text: 'BIG TEXT HERE' },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      } satisfies Table,
    ]
    expect(layoutDocument(tight).pages[0]!.blocks[0]?.kind).toBe('table')
  })
})

