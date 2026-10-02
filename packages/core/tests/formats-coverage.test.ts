import { describe, expect, it } from 'vitest'
import { strToU8, zipSync } from 'fflate'
import {
  AlmadocxError,
  createEmptyDocument,
  loadDocument,
  parseDocx,
  parseOdt,
  saveDocument,
  serializeDocx,
  serializeOdt,
  sniffDocx,
  sniffOdt,
  type Document,
} from '../src/index.js'
import { parseNumberingXml, serializeNumberingXml } from '../src/formats/docx/numbering.js'
import { setZipText, writeZip } from '../src/io/zip.js'

function richDocxDoc(): Document {
  const doc = createEmptyDocument('docx')
  doc.media = {
    png: { id: 'png', contentType: 'image/png', bytes: new Uint8Array([137, 80, 78, 71]) },
    jpg: { id: 'jpg', contentType: 'image/jpeg', bytes: new Uint8Array([1, 2]) },
    gif: { id: 'gif', contentType: 'image/gif', bytes: new Uint8Array([1]) },
    webp: { id: 'webp', contentType: 'image/webp', bytes: new Uint8Array([1]) },
  }
  doc.numbering.nums['3'] = {
    numId: '3',
    abstractNumId: 'abs_decimal',
    startOverrides: { 0: 10 },
  }
  doc.package.preservedParts = [
    { path: 'customXml/item1.xml', bytes: strToU8('<x/>') },
    { path: 'word/document.xml', bytes: strToU8('skip') },
    { path: 'word/media/old.png', bytes: new Uint8Array([9]) },
  ]
  doc.sections[0]!.header = {
    blocks: [
      {
        id: 'h',
        type: 'paragraph',
        props: {},
        runs: [{ id: 'hr', props: {}, content: { type: 'text', text: 'H' } }],
      },
    ],
  }
  doc.sections[0]!.footer = {
    blocks: [
      {
        id: 'f',
        type: 'paragraph',
        props: {},
        runs: [{ id: 'fr', props: {}, content: { type: 'text', text: 'F' } }],
      },
    ],
  }
  doc.sections[0]!.blocks = [
    {
      id: 'p1',
      type: 'paragraph',
      props: {
        styleId: 'Normal',
        numPr: { numId: '3', ilvl: 0 },
        alignment: 'justify',
        indentLeft: 200,
        indentRight: 100,
        indentFirstLine: -120,
        spacingBefore: 100,
        spacingAfter: 100,
        lineSpacing: 1.5,
        lineSpacingRule: 'auto',
        keepNext: true,
        keepLines: true,
        widowControl: false,
        pageBreakBefore: true,
      },
      runs: [
        {
          id: 'r1',
          props: {
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            fontSizePt: 14,
            color: '#112233',
            fontFamily: 'Arial',
            verticalAlign: 'superscript',
          },
          content: { type: 'text', text: ' Styled ' },
        },
        { id: 'r2', props: {}, content: { type: 'tab' } },
        { id: 'r3', props: {}, content: { type: 'break', breakType: 'page' } },
        { id: 'r4', props: {}, content: { type: 'break', breakType: 'line' } },
        {
          id: 'r5',
          props: {},
          content: { type: 'image', mediaId: 'png', widthTwips: 500, heightTwips: 500, alt: 'a' },
        },
        {
          id: 'r6',
          props: {},
          content: { type: 'image', mediaId: 'missing', widthTwips: 10, heightTwips: 10 },
        },
        { id: 'r7', props: { verticalAlign: 'subscript' }, content: { type: 'text', text: 'sub' } },
        { id: 'r8', props: { color: 'nope' }, content: { type: 'text', text: 'badcolor' } },
      ],
    },
    {
      id: 'p2',
      type: 'paragraph',
      props: {
        indentFirstLine: 240,
        lineSpacing: 480,
        lineSpacingRule: 'exact',
      },
      runs: [{ id: 'plain', props: {}, content: { type: 'text', text: 'plain' } }],
    },
    {
      id: 'tbl',
      type: 'table',
      props: {
        widthTwips: 5000,
        alignment: 'center',
        cellSpacing: 40,
        borders: {
          top: '#000000',
          bottom: '#111111',
          left: '#222222',
          right: '#333333',
          insideH: '#444444',
          insideV: '#555555',
        },
      },
      gridCols: [2500, 2500],
      rows: [
        {
          id: 'hr',
          props: { header: true, heightTwips: 400 },
          cells: [
            {
              id: 'c0',
              props: {
                widthTwips: 2500,
                gridSpan: 1,
                vMerge: 'restart',
                shading: '#abcdef',
                vAlign: 'center',
                borders: { top: '#aa0000', bottom: '#00aa00', left: '#0000aa', right: '#aaaa00' },
                margin: { top: 10, left: 10, bottom: 10, right: 10 },
              },
              blocks: [
                {
                  id: 'cp0',
                  type: 'paragraph',
                  props: {},
                  runs: [{ id: 'cr0', props: {}, content: { type: 'text', text: 'A' } }],
                },
              ],
            },
            {
              id: 'c1',
              props: { shading: 'zzzzzz' },
              blocks: [
                {
                  id: 'cp1',
                  type: 'paragraph',
                  props: {},
                  runs: [{ id: 'cr1', props: {}, content: { type: 'text', text: 'B' } }],
                },
              ],
            },
          ],
        },
        {
          id: 'r1',
          props: {},
          cells: [
            {
              id: 'c2',
              props: { vMerge: 'continue' },
              blocks: [
                {
                  id: 'cp2',
                  type: 'paragraph',
                  props: {},
                  runs: [{ id: 'cr2', props: {}, content: { type: 'text', text: '' } }],
                },
              ],
            },
            {
              id: 'c3',
              props: {},
              blocks: [
                {
                  id: 'cp3',
                  type: 'paragraph',
                  props: {},
                  runs: [{ id: 'cr3', props: {}, content: { type: 'text', text: 'D' } }],
                },
              ],
            },
          ],
        },
      ],
    },
  ]
  return doc
}

describe('docx serialize/parse + numbering', () => {
  it('serializes rich docs and round-trips', () => {
    const doc = richDocxDoc()
    const bytes = serializeDocx(doc)
    expect(sniffDocx(bytes)).toBe(true)
    const loaded = parseDocx(bytes)
    expect(loaded.sections[0]!.blocks.length).toBeGreaterThan(1)
    expect(loaded.sections[0]!.header).toBeTruthy()
    expect(loaded.sections[0]!.footer).toBeTruthy()
    expect(Object.keys(loaded.media).length).toBeGreaterThan(0)

    const again = saveDocument(loaded, 'docx')
    expect(loadDocument(again, 'docx').sections.length).toBe(1)
  })

  it('parses numbering formats and overrides', () => {
    expect(parseNumberingXml(undefined).nums['1']).toBeTruthy()
    const xml = serializeNumberingXml({
      abstractNums: {
        a1: {
          id: 'a1',
          levels: [
            {
              ilvl: 0,
              format: 'lowerRoman',
              levelText: '%1.',
              start: 1,
              indentLeft: 720,
              hanging: 360,
              fontFamily: 'Symbol',
            },
            { ilvl: 1, format: 'upperLetter', levelText: '%2)', start: 1 },
            { ilvl: 2, format: 'none', levelText: '', start: 1 },
            { ilvl: 3, format: 'bullet', levelText: '•', start: 1 },
            { ilvl: 4, format: 'lowerLetter', levelText: '%5.', start: 1 },
            { ilvl: 5, format: 'decimal', levelText: '%6.', start: 1 },
          ],
        },
      },
      nums: {
        '5': { numId: '5', abstractNumId: 'a1', startOverrides: { 0: 3 } },
      },
    })
    expect(xml).toContain('lvlOverride')
    const parsed = parseNumberingXml(xml)
    expect(parsed.nums['5']?.startOverrides?.[0]).toBe(3)

    // mapNumFmt default / unknown
    const odd = parseNumberingXml(
      `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:abstractNum w:abstractNumId="z"><w:lvl w:ilvl="0">` +
        `<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1."/><w:start w:val="1"/>` +
        `</w:lvl></w:abstractNum>` +
        `<w:num w:numId="8"><w:abstractNumId w:val="z"/></w:num>` +
        `</w:numbering>`,
    )
    expect(odd.abstractNums['z']?.levels[0]?.format).toBe('decimal')

    const emptyFmt = parseNumberingXml(
      `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:abstractNum w:abstractNumId="b"><w:lvl w:ilvl="0"><w:lvlText w:val="x"/></w:lvl></w:abstractNum>` +
        `</w:numbering>`,
    )
    expect(emptyFmt.abstractNums['b']?.levels[0]?.format).toBe('bullet')
  })
})

describe('odt serialize/parse', () => {
  it('serializes styled runs/paras and parses content', () => {
    let doc = createEmptyDocument('odt')
    doc.package.preservedParts = [{ path: 'meta.xml', bytes: strToU8('<meta/>') }]
    doc.sections[0]!.blocks = [
      {
        id: 'p1',
        type: 'paragraph',
        props: {
          alignment: 'center',
          indentLeft: 200,
          indentRight: 100,
          indentFirstLine: 50,
          spacingBefore: 40,
          spacingAfter: 40,
          lineSpacing: 1.2,
          lineSpacingRule: 'auto',
        },
        runs: [
          {
            id: 'r1',
            props: {
              bold: true,
              italic: true,
              underline: true,
              strike: true,
              fontSizePt: 16,
              color: '#abc',
              fontFamily: 'DejaVu Sans',
            },
            content: { type: 'text', text: 'Hi &<>"' },
          },
          { id: 'r2', props: {}, content: { type: 'tab' } },
          { id: 'r3', props: {}, content: { type: 'break', breakType: 'line' } },
          {
            id: 'r4',
            props: {},
            content: { type: 'image', mediaId: 'x', widthTwips: 10, heightTwips: 10, alt: 'pic' },
          },
          { id: 'r5', props: {}, content: { type: 'text', text: 'plain' } },
        ],
      },
    ]
    const bytes = serializeOdt(doc)
    expect(sniffOdt(bytes)).toBe(true)
    const loaded = parseOdt(bytes)
    expect(loaded.sections[0]!.blocks.length).toBeGreaterThan(0)

    // Hand-built ODT with styles + table for parse branches
    const entries = new Map<string, Uint8Array>()
    setZipText(entries, 'mimetype', 'application/vnd.oasis.opendocument.text')
    setZipText(
      entries,
      'styles.xml',
      `<?xml version="1.0"?>` +
        `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
        `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">` +
        `<office:styles>` +
        `<style:style style:name="PCenter" style:family="paragraph">` +
        `<style:paragraph-properties fo:text-align="center" fo:margin-left="1cm" fo:margin-right="2mm" ` +
        `fo:margin-top="0.1in" fo:margin-bottom="12pt" fo:text-indent="6pc" fo:line-height="150%"/>` +
        `<style:text-properties fo:font-weight="bold" fo:font-style="italic" style:text-underline-style="solid" ` +
        `style:text-line-through-style="solid" fo:font-size="12pt" fo:color="#00ff00" style:font-name="Arial"/>` +
        `</style:style>` +
        `<style:style style:name="TBold" style:family="text">` +
        `<style:text-properties fo:font-weight="700" fo:font-family="Helvetica"/>` +
        `</style:style>` +
        `<style:style style:name="PLeft" style:family="paragraph">` +
        `<style:paragraph-properties fo:text-align="start"/>` +
        `</style:style>` +
        `<style:style style:name="PRight" style:family="paragraph">` +
        `<style:paragraph-properties text-align="end"/>` +
        `</style:style>` +
        `<style:style style:name="PJust" style:family="paragraph">` +
        `<style:paragraph-properties fo:text-align="justify"/>` +
        `</style:style>` +
        `</office:styles></office:document-styles>`,
    )
    setZipText(
      entries,
      'content.xml',
      `<?xml version="1.0"?>` +
        `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">` +
        `<office:body><office:text>` +
        `<text:p text:style-name="PCenter">Hello <text:span text:style-name="TBold">Bold</text:span>` +
        `<text:tab/><text:line-break/><text:s text:c="2"/>world</text:p>` +
        `<text:p text:style-name="PLeft">L</text:p>` +
        `<text:p text:style-name="PRight">R</text:p>` +
        `<text:p text:style-name="PJust">J</text:p>` +
        `<table:table>` +
        `<table:table-column/><table:table-column/>` +
        `<table:table-row>` +
        `<table:table-cell><text:p>A</text:p></table:table-cell>` +
        `<table:table-cell table:number-columns-spanned="1"><text:p>B</text:p></table:table-cell>` +
        `</table:table-row>` +
        `</table:table>` +
        `<text:list><text:list-item><text:p>item</text:p></text:list-item></text:list>` +
        `</office:text></office:body></office:document-content>`,
    )
    setZipText(
      entries,
      'META-INF/manifest.xml',
      `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`,
    )
    // macro path for filter
    setZipText(entries, 'Basic/Module.xml', '<module/>')
    const odtBytes = writeZip(entries)
    const odt = parseOdt(odtBytes)
    expect(odt.sections[0]!.blocks.length).toBeGreaterThan(2)
    expect(odt.package.macroPartPaths.length).toBeGreaterThan(0)
  })

  it('sniff helpers reject non-packages', () => {
    expect(sniffDocx(new Uint8Array([0, 1]))).toBe(false)
    expect(sniffOdt(new Uint8Array([0, 1]))).toBe(false)
    expect(sniffOdt(zipSync({ 'other.txt': strToU8('x') }))).toBe(false)
  })
})

describe('docx parse edge XML', () => {
  it('parses document with styles, sectPr, drawings, and on/off attrs', () => {
    const entries = new Map<string, Uint8Array>()
    setZipText(
      entries,
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `</Types>`,
    )
    setZipText(
      entries,
      '_rels/.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
        `</Relationships>`,
    )
    setZipText(
      entries,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/>` +
        `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
        `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>` +
        `</Relationships>`,
    )
    setZipText(entries, 'word/media/img.png', '\x89PNG')
    setZipText(
      entries,
      'word/styles.xml',
      `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:docDefaults><w:rPrDefault><w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="FF0000"/></w:rPr></w:rPrDefault>` +
        `<w:pPrDefault><w:pPr><w:jc w:val="center"/></w:pPr></w:pPrDefault></w:docDefaults>` +
        `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>` +
        `<w:pPr><w:jc w:val="both"/><w:ind w:left="100" w:right="50" w:firstLine="20"/><w:spacing w:before="10" w:after="10" w:line="480" w:lineRule="exact"/></w:pPr>` +
        `<w:rPr><w:i w:val="1"/><w:u w:val="single"/><w:strike w:val="true"/><w:vertAlign w:val="superscript"/></w:rPr></w:style>` +
        `<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b w:val="0"/><w:highlight w:val="yellow"/></w:rPr></w:style>` +
        `</w:styles>`,
    )
    setZipText(
      entries,
      'word/numbering.xml',
      serializeNumberingXml(createEmptyDocument().numbering),
    )
    setZipText(
      entries,
      'word/header1.xml',
      `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>`,
    )
    setZipText(
      entries,
      'word/footer1.xml',
      `<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>F</w:t></w:r></w:p></w:ftr>`,
    )
    setZipText(
      entries,
      'word/document.xml',
      `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<w:body>` +
        `<w:p><w:pPr><w:pStyle w:val="Title"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` +
        `<w:jc w:val="left"/><w:ind w:hanging="120"/><w:spacing w:line="276" w:lineRule="auto"/>` +
        `<w:keepNext/><w:keepLines/><w:widowControl w:val="0"/><w:pageBreakBefore/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/><w:sz w:val="28"/><w:color w:val="00FF00"/>` +
        `<w:rFonts w:ascii="Arial"/><w:vertAlign w:val="subscript"/><w:highlight w:val="green"/></w:rPr>` +
        `<w:t xml:space="preserve"> Hi </w:t></w:r>` +
        `<w:r><w:tab/></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:br/></w:r>` +
        `<w:r><w:drawing><wp:inline><wp:extent cx="635000" cy="635000"/>` +
        `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId3"/></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:ext cx="635000" cy="635000"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>` +
        `</wp:inline></w:drawing></w:r>` +
        `</w:p>` +
        `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:jc w:val="center"/>` +
        `<w:tblBorders><w:top w:val="single" w:color="000000"/><w:bottom w:val="single" w:color="111111"/>` +
        `<w:left w:val="single" w:color="222222"/><w:right w:val="single" w:color="333333"/>` +
        `<w:insideH w:val="single" w:color="444444"/><w:insideV w:val="single" w:color="555555"/></w:tblBorders>` +
        `<w:tblCellSpacing w:w="20" w:type="dxa"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="300"/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="2500"/><w:gridSpan w:val="1"/><w:vMerge/>` +
        `<w:tcBorders><w:top w:val="single" w:color="AA0000"/></w:tcBorders>` +
        `<w:shd w:fill="ABCDEF"/><w:vAlign w:val="center"/>` +
        `<w:tcMar><w:top w:w="10" w:type="dxa"/><w:left w:w="10" w:type="dxa"/></w:tcMar></w:tcPr>` +
        `<w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>` +
        `</w:tr></w:tbl>` +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>` +
        `<w:cols w:num="2"/><w:titlePg/><w:headerReference w:type="default" r:id="rId4"/>` +
        `<w:footerReference w:type="default" r:id="rId5"/></w:sectPr>` +
        `</w:body></w:document>`,
    )
    // macro-ish preserved part
    setZipText(entries, 'word/vbaProject.bin', 'MZ')

    const parsed = parseDocx(writeZip(entries))
    expect(parsed.sections[0]!.properties.columns).toBe(2)
    expect(parsed.sections[0]!.header).toBeTruthy()
    expect(parsed.styles.paragraphStyles['Title']).toBeTruthy()
    expect(parsed.package.macroPartPaths.some((p) => p.includes('vba'))).toBe(true)
  })

  it('rejects packages without document.xml', () => {
    const entries = new Map<string, Uint8Array>()
    setZipText(
      entries,
      '[Content_Types].xml',
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    )
    expect(() => parseDocx(writeZip(entries))).toThrow(AlmadocxError)
  })
})
