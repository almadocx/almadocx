import type {
  Block,
  CharacterProps,
  Document,
  HeaderFooter,
  Paragraph,
  ParagraphProps,
  Run,
  Table,
} from '../../model/types.js'
import { setZipText, writeZip, type ZipEntries } from '../../io/zip.js'
import { AlmadocxError } from '../../util/assert.js'
import { serializeNumberingXml } from './numbering.js'

function esc(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function colorToHex(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#?([0-9A-Fa-f]{6})$/.exec(color)
  return m?.[1]?.toUpperCase()
}

function rPrXml(props: CharacterProps): string {
  const parts: string[] = []
  if (props.bold) parts.push('<w:b/>')
  if (props.italic) parts.push('<w:i/>')
  if (props.underline) parts.push('<w:u w:val="single"/>')
  if (props.strike) parts.push('<w:strike/>')
  if (props.fontSizePt !== undefined) {
    parts.push(`<w:sz w:val="${String(Math.round(props.fontSizePt * 2))}"/>`)
    parts.push(`<w:szCs w:val="${String(Math.round(props.fontSizePt * 2))}"/>`)
  }
  const hex = colorToHex(props.color)
  if (hex) parts.push(`<w:color w:val="${hex}"/>`)
  if (props.fontFamily) {
    const f = esc(props.fontFamily)
    parts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`)
  }
  if (props.verticalAlign === 'superscript' || props.verticalAlign === 'subscript') {
    parts.push(`<w:vertAlign w:val="${props.verticalAlign}"/>`)
  }
  if (parts.length === 0) return ''
  return `<w:rPr>${parts.join('')}</w:rPr>`
}

function pPrXml(props: ParagraphProps): string {
  const parts: string[] = []
  if (props.styleId) parts.push(`<w:pStyle w:val="${esc(props.styleId)}"/>`)
  if (props.numPr) {
    parts.push(
      `<w:numPr><w:ilvl w:val="${String(props.numPr.ilvl)}"/><w:numId w:val="${esc(props.numPr.numId)}"/></w:numPr>`,
    )
  }
  if (props.alignment) {
    const jc = props.alignment === 'justify' ? 'both' : props.alignment
    parts.push(`<w:jc w:val="${jc}"/>`)
  }
  const indAttrs: string[] = []
  if (props.indentLeft !== undefined) indAttrs.push(`w:left="${String(props.indentLeft)}"`)
  if (props.indentRight !== undefined) indAttrs.push(`w:right="${String(props.indentRight)}"`)
  if (props.indentFirstLine !== undefined) {
    if (props.indentFirstLine < 0) {
      indAttrs.push(`w:hanging="${String(-props.indentFirstLine)}"`)
    } else {
      indAttrs.push(`w:firstLine="${String(props.indentFirstLine)}"`)
    }
  }
  if (indAttrs.length) parts.push(`<w:ind ${indAttrs.join(' ')}/>`)
  const sp: string[] = []
  if (props.spacingBefore !== undefined) sp.push(`w:before="${String(props.spacingBefore)}"`)
  if (props.spacingAfter !== undefined) sp.push(`w:after="${String(props.spacingAfter)}"`)
  if (props.lineSpacing !== undefined) {
    if (props.lineSpacingRule === 'auto' || props.lineSpacingRule === undefined) {
      sp.push(`w:line="${String(Math.round(props.lineSpacing * 240))}"`)
      sp.push('w:lineRule="auto"')
    } else {
      sp.push(`w:line="${String(Math.round(props.lineSpacing))}"`)
      sp.push(`w:lineRule="${props.lineSpacingRule}"`)
    }
  }
  if (sp.length) parts.push(`<w:spacing ${sp.join(' ')}/>`)
  if (props.keepNext) parts.push('<w:keepNext/>')
  if (props.keepLines) parts.push('<w:keepLines/>')
  if (props.widowControl === false) parts.push('<w:widowControl w:val="0"/>')
  if (props.pageBreakBefore) parts.push('<w:pageBreakBefore/>')
  if (parts.length === 0) return ''
  return `<w:pPr>${parts.join('')}</w:pPr>`
}

function runXml(run: Run, imageRels: Map<string, string>): string {
  const pr = rPrXml(run.props)
  switch (run.content.type) {
    case 'text': {
      const space = /^\s|\s$/.test(run.content.text) ? ' xml:space="preserve"' : ''
      return `<w:r>${pr}<w:t${space}>${esc(run.content.text)}</w:t></w:r>`
    }
    case 'tab':
      return `<w:r>${pr}<w:tab/></w:r>`
    case 'break': {
      const type = run.content.breakType === 'line' ? '' : ` w:type="${run.content.breakType}"`
      return `<w:r>${pr}<w:br${type}/></w:r>`
    }
    case 'image': {
      const rid = imageRels.get(run.content.mediaId)
      if (!rid) return `<w:r>${pr}</w:r>`
      const cx = run.content.widthTwips * 635
      const cy = run.content.heightTwips * 635
      return (
        `<w:r>${pr}<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
        `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
        `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<wp:extent cx="${String(cx)}" cy="${String(cy)}"/>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
        `<pic:pic><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:ext cx="${String(cx)}" cy="${String(cy)}"/></a:xfrm></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
      )
    }
  }
}

function paragraphXml(p: Paragraph, imageRels: Map<string, string>): string {
  return `<w:p>${pPrXml(p.props)}${p.runs.map((r) => runXml(r, imageRels)).join('')}</w:p>`
}

function bordersXml(
  tag: 'w:tblBorders' | 'w:tcBorders',
  borders: { [key: string]: string | undefined } | undefined,
  edges: readonly string[],
): string {
  if (!borders) return ''
  const parts: string[] = []
  for (const edge of edges) {
    const color = colorToHex(borders[edge])
    if (!color) continue
    parts.push(`<w:${edge} w:val="single" w:sz="4" w:space="0" w:color="${color}"/>`)
  }
  if (!parts.length) return ''
  return `<${tag}>${parts.join('')}</${tag}>`
}

function tblPrXml(table: Table): string {
  const parts: string[] = []
  if (table.props.widthTwips !== undefined) {
    parts.push(`<w:tblW w:w="${String(table.props.widthTwips)}" w:type="dxa"/>`)
  }
  if (table.props.alignment) {
    parts.push(`<w:jc w:val="${table.props.alignment}"/>`)
  }
  parts.push(
    bordersXml('w:tblBorders', table.props.borders as { [key: string]: string | undefined } | undefined, [
      'top',
      'bottom',
      'left',
      'right',
      'insideH',
      'insideV',
    ]),
  )
  if (table.props.cellSpacing !== undefined) {
    parts.push(`<w:tblCellSpacing w:w="${String(table.props.cellSpacing)}" w:type="dxa"/>`)
  }
  const inner = parts.filter(Boolean).join('')
  return inner ? `<w:tblPr>${inner}</w:tblPr>` : ''
}

function tableXml(table: Table, imageRels: Map<string, string>): string {
  const grid =
    table.gridCols?.map((w) => `<w:gridCol w:w="${String(w)}"/>`).join('') ??
    `<w:gridCol w:w="2400"/><w:gridCol w:w="2400"/>`
  const rows = table.rows
    .map((row) => {
      const cells = row.cells
        .map((cell) => {
          const prParts: string[] = []
          if (cell.props.widthTwips !== undefined) {
            prParts.push(`<w:tcW w:w="${String(cell.props.widthTwips)}" w:type="dxa"/>`)
          }
          if (cell.props.gridSpan && cell.props.gridSpan > 1) {
            prParts.push(`<w:gridSpan w:val="${String(cell.props.gridSpan)}"/>`)
          }
          if (cell.props.vMerge) {
            prParts.push(
              cell.props.vMerge === 'continue' ? '<w:vMerge w:val="continue"/>' : '<w:vMerge/>',
            )
          }
          prParts.push(
            bordersXml('w:tcBorders', cell.props.borders as { [key: string]: string | undefined } | undefined, [
              'top',
              'bottom',
              'left',
              'right',
            ]),
          )
          if (cell.props.shading) {
            const fill = colorToHex(cell.props.shading) ?? 'FFFFFF'
            prParts.push(`<w:shd w:val="clear" w:fill="${fill}"/>`)
          }
          if (cell.props.margin) {
            const m = cell.props.margin
            const edges: string[] = []
            for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
              const v = m[edge]
              if (v !== undefined) edges.push(`<w:${edge} w:w="${String(v)}" w:type="dxa"/>`)
            }
            if (edges.length) prParts.push(`<w:tcMar>${edges.join('')}</w:tcMar>`)
          }
          if (cell.props.vAlign) {
            prParts.push(`<w:vAlign w:val="${cell.props.vAlign}"/>`)
          }
          const tcPrInner = prParts.filter(Boolean).join('')
          const tcPr = tcPrInner ? `<w:tcPr>${tcPrInner}</w:tcPr>` : ''
          const body = cell.blocks.map((p) => paragraphXml(p, imageRels)).join('')
          return `<w:tc>${tcPr}${body}</w:tc>`
        })
        .join('')
      const trPrParts: string[] = []
      if (row.props.header) trPrParts.push('<w:tblHeader/>')
      if (row.props.heightTwips !== undefined) {
        trPrParts.push(`<w:trHeight w:val="${String(row.props.heightTwips)}"/>`)
      }
      const trPr = trPrParts.length ? `<w:trPr>${trPrParts.join('')}</w:trPr>` : ''
      return `<w:tr>${trPr}${cells}</w:tr>`
    })
    .join('')
  return `<w:tbl>${tblPrXml(table)}<w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`
}

function blockXml(block: Block, imageRels: Map<string, string>): string {
  if (block.type === 'paragraph') return paragraphXml(block, imageRels)
  return tableXml(block, imageRels)
}

function headerFooterXml(hf: HeaderFooter, imageRels: Map<string, string>, tag: 'hdr' | 'ftr'): string {
  const body = hf.blocks.map((p) => paragraphXml(p, imageRels)).join('')
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    body +
    `</w:${tag}>`
  )
}

function buildDocumentXml(doc: Document, imageRels: Map<string, string>): string {
  const body = doc.sections.flatMap((s) => s.blocks).map((b) => blockXml(b, imageRels)).join('')
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<w:body>${body}<w:sectPr/></w:body></w:document>`
  )
}

export function serializeDocx(doc: Document): Uint8Array {
  const entries: ZipEntries = new Map()
  for (const part of doc.package.preservedParts) {
    if (
      part.path === 'word/document.xml' ||
      part.path === 'word/numbering.xml' ||
      part.path.startsWith('word/media/')
    ) {
      continue
    }
    entries.set(part.path, part.bytes)
  }

  const imageRels = new Map<string, string>()
  const relParts: string[] = [
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`,
  ]
  let rid = 3
  const contentOverrides: string[] = [
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
    `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`,
  ]

  for (const [id, media] of Object.entries(doc.media)) {
    const ext =
      media.contentType === 'image/jpeg'
        ? 'jpeg'
        : media.contentType === 'image/gif'
          ? 'gif'
          : media.contentType === 'image/webp'
            ? 'webp'
            : 'png'
    const path = media.path ?? `word/media/${id}.${ext}`
    entries.set(path, media.bytes)
    const relId = `rId${String(rid++)}`
    imageRels.set(id, relId)
    const target = path.replace(/^word\//, '')
    relParts.push(
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${esc(target)}"/>`,
    )
  }

  const section = doc.sections[0]
  if (section?.header) {
    setZipText(entries, 'word/header1.xml', headerFooterXml(section.header, imageRels, 'hdr'))
    const relId = `rId${String(rid++)}`
    relParts.push(
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`,
    )
    contentOverrides.push(
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
    )
  }
  if (section?.footer) {
    setZipText(entries, 'word/footer1.xml', headerFooterXml(section.footer, imageRels, 'ftr'))
    const relId = `rId${String(rid++)}`
    relParts.push(
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`,
    )
    contentOverrides.push(
      `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`,
    )
  }

  setZipText(entries, 'word/document.xml', buildDocumentXml(doc, imageRels))
  setZipText(entries, 'word/numbering.xml', serializeNumberingXml(doc.numbering))

  if (!entries.has('word/styles.xml')) {
    setZipText(
      entries,
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>` +
        `<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
        `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
        `</w:styles>`,
    )
  }

  setZipText(
    entries,
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      relParts.join('') +
      `</Relationships>`,
  )

  if (!entries.has('_rels/.rels')) {
    setZipText(
      entries,
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
        `</Relationships>`,
    )
  }

  setZipText(
    entries,
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
      `<Default Extension="jpg" ContentType="image/jpeg"/>` +
      contentOverrides.join('') +
      `</Types>`,
  )

  try {
    return writeZip(entries)
  } catch (e) {
    throw new AlmadocxError('docx_serialize', `failed to write docx: ${String(e)}`)
  }
}
