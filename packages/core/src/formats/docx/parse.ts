import {
  createEmptyDocument,
  DEFAULT_MARGINS,
  DEFAULT_PAGE_SIZE,
  type Block,
  type CharacterProps,
  type Document,
  type HeaderFooter,
  type Paragraph,
  type ParagraphProps,
  type Run,
  type Section,
  type SectionProperties,
  type Table,
  type TableCell,
  type TableRow,
} from '../../model/types.js'
import { createTextRun } from '../../model/text.js'
import { isMacroPath, readZip, zipText, type ZipEntries } from '../../io/zip.js'
import { asArray, parseXml, xmlAttrAlt, xmlText, xmlAttrAltOr } from '../../io/xml.js'
import { nextId } from '../../util/id.js'
import { assert } from '../../util/assert.js'
import { parseNumberingXml } from './numbering.js'

function local(name: string): string {
  const i = name.indexOf(':')
  return i >= 0 ? name.slice(i + 1) : name
}

function findChild(node: Record<string, unknown>, localName: string): unknown {
  for (const [k, v] of Object.entries(node)) {
    if (local(k) === localName) return v
  }
  return undefined
}

function findChildren(node: Record<string, unknown>, localName: string): unknown[] {
  const v = findChild(node, localName)
  return asArray(v as never)
}

function parseOnOff(val: string | undefined): boolean | undefined {
  if (val === undefined) return true
  if (val === '0' || val === 'false' || val === 'off') return false
  return true
}

function parseColor(val: string | undefined): string | undefined {
  if (!val || val === 'auto') return undefined
  if (/^[0-9A-Fa-f]{6}$/.test(val)) return `#${val}`
  return undefined
}

function parseRunProps(rPr: Record<string, unknown> | undefined): CharacterProps {
  if (!rPr) return {}
  const props: CharacterProps = {}
  if (findChild(rPr, 'b') !== undefined) {
    const b = findChild(rPr, 'b') as Record<string, unknown> | string
    const val = typeof b === 'object' ? xmlAttrAlt(b, 'w:val', 'val') : undefined
    const on = parseOnOff(val)
    if (on !== undefined) props.bold = on
  }
  if (findChild(rPr, 'i') !== undefined) {
    const i = findChild(rPr, 'i') as Record<string, unknown> | string
    const val = typeof i === 'object' ? xmlAttrAlt(i, 'w:val', 'val') : undefined
    const on = parseOnOff(val)
    if (on !== undefined) props.italic = on
  }
  if (findChild(rPr, 'u') !== undefined) {
    const u = findChild(rPr, 'u') as Record<string, unknown>
    const val = xmlAttrAlt(u, 'w:val', 'val')
    props.underline = val !== 'none'
  }
  if (findChild(rPr, 'strike') !== undefined) props.strike = true
  const sz = findChild(rPr, 'sz') as Record<string, unknown> | undefined
  if (sz) {
    const v = xmlAttrAlt(sz, 'w:val', 'val')
    if (v) props.fontSizePt = Number(v) / 2
  }
  const color = findChild(rPr, 'color') as Record<string, unknown> | undefined
  if (color) {
    const c = parseColor(xmlAttrAlt(color, 'w:val', 'val'))
    if (c) props.color = c
  }
  const rFonts = findChild(rPr, 'rFonts') as Record<string, unknown> | undefined
  if (rFonts) {
    const ascii =
      xmlAttrAlt(rFonts, 'w:ascii', 'ascii', 'w:hAnsi', 'hAnsi')
    if (ascii) props.fontFamily = ascii
  }
  const vert = findChild(rPr, 'vertAlign') as Record<string, unknown> | undefined
  if (vert) {
    const v = xmlAttrAlt(vert, 'w:val', 'val')
    if (v === 'superscript' || v === 'subscript') props.verticalAlign = v
  }
  return props
}

function parseParagraphProps(pPr: Record<string, unknown> | undefined): ParagraphProps {
  if (!pPr) return {}
  const props: ParagraphProps = {}
  const pStyle = findChild(pPr, 'pStyle') as Record<string, unknown> | undefined
  if (pStyle) {
    const id = xmlAttrAlt(pStyle, 'w:val', 'val')
    if (id) props.styleId = id
  }
  const jc = findChild(pPr, 'jc') as Record<string, unknown> | undefined
  if (jc) {
    const v = xmlAttrAlt(jc, 'w:val', 'val')
    if (v === 'left' || v === 'center' || v === 'right' || v === 'both') {
      props.alignment = v === 'both' ? 'justify' : v
    }
  }
  const ind = findChild(pPr, 'ind') as Record<string, unknown> | undefined
  if (ind) {
    const left = xmlAttrAlt(ind, 'w:left', 'left')
    const right = xmlAttrAlt(ind, 'w:right', 'right')
    const first = xmlAttrAlt(ind, 'w:firstLine', 'firstLine')
    const hanging = xmlAttrAlt(ind, 'w:hanging', 'hanging')
    if (left) props.indentLeft = Number(left)
    if (right) props.indentRight = Number(right)
    if (first) props.indentFirstLine = Number(first)
    if (hanging) props.indentFirstLine = -Number(hanging)
  }
  const spacing = findChild(pPr, 'spacing') as Record<string, unknown> | undefined
  if (spacing) {
    const before = xmlAttrAlt(spacing, 'w:before', 'before')
    const after = xmlAttrAlt(spacing, 'w:after', 'after')
    const line = xmlAttrAlt(spacing, 'w:line', 'line')
    const lineRule = xmlAttrAlt(spacing, 'w:lineRule', 'lineRule')
    if (before) props.spacingBefore = Number(before)
    if (after) props.spacingAfter = Number(after)
    if (line) {
      if (lineRule === 'auto' || !lineRule) {
        props.lineSpacing = Number(line) / 240
        props.lineSpacingRule = 'auto'
      } else if (lineRule === 'exact') {
        props.lineSpacing = Number(line)
        props.lineSpacingRule = 'exact'
      } else {
        props.lineSpacing = Number(line)
        props.lineSpacingRule = 'atLeast'
      }
    }
  }
  const numPr = findChild(pPr, 'numPr') as Record<string, unknown> | undefined
  if (numPr) {
    const ilvlNode = findChild(numPr, 'ilvl') as Record<string, unknown> | undefined
    const numIdNode = findChild(numPr, 'numId') as Record<string, unknown> | undefined
    const ilvl = Number(xmlAttrAltOr(ilvlNode, '0', 'w:val', 'val'))
    const numId = xmlAttrAlt(numIdNode, 'w:val', 'val')
    if (numId) props.numPr = { numId, ilvl }
  }
  if (findChild(pPr, 'keepNext') !== undefined) props.keepNext = true
  if (findChild(pPr, 'keepLines') !== undefined) props.keepLines = true
  const widow = findChild(pPr, 'widowControl') as Record<string, unknown> | undefined
  if (widow) {
    const val = xmlAttrAlt(widow, 'w:val', 'val')
    props.widowControl = val !== '0' && val !== 'false'
  }
  if (findChild(pPr, 'pageBreakBefore') !== undefined) props.pageBreakBefore = true
  return props
}

function parseRun(rNode: Record<string, unknown>, doc?: Document, rels?: Map<string, string>): Run[] {
  const rPr = findChild(rNode, 'rPr') as Record<string, unknown> | undefined
  const props = parseRunProps(rPr)
  const runs: Run[] = []
  let pageField = false

  for (const [key, value] of Object.entries(rNode)) {
    const name = local(key)
    if (name === 'rPr' || key.startsWith('@_')) continue
    if (name === 't') {
      const text = xmlText(value)
      runs.push(createTextRun(text, props))
    } else if (name === 'tab') {
      runs.push({ id: nextId('r'), props, content: { type: 'tab' } })
    } else if (name === 'br') {
      const br = value as Record<string, unknown>
      const type = xmlAttrAlt(br, 'w:type', 'type')
      const breakType = type === 'page' ? 'page' : type === 'column' ? 'column' : 'line'
      runs.push({ id: nextId('r'), props, content: { type: 'break', breakType } })
    } else if (name === 'drawing' && doc && rels) {
      const img = parseDrawingImage(value, doc, rels)
      if (img) runs.push({ id: nextId('r'), props, content: img })
    } else if (name === 'instrText') {
      const instr = xmlText(value).trim().toUpperCase()
      // Complex fields: PAGE / NUMPAGES etc.
      if (instr === 'PAGE' || instr.startsWith('PAGE ') || instr.startsWith('PAGE\\')) {
        pageField = true
      }
    }
  }

  if (pageField) {
    runs.push(createTextRun('{{PAGE}}', props))
  }

  if (runs.length === 0) runs.push(createTextRun('', props))
  return runs
}

function parseDrawingImage(
  drawing: unknown,
  doc: Document,
  rels: Map<string, string>,
): Extract<Run['content'], { type: 'image' }> | undefined {
  const nodes = asArray(drawing as never)
  for (const d of nodes) {
    const blob = JSON.stringify(d)
    const embedMatch = /"@_r:embed":"([^"]+)"/.exec(blob) ?? /"@_embed":"([^"]+)"/.exec(blob)
    const cxMatch = /"@_cx":"(\d+)"/.exec(blob)
    const cyMatch = /"@_cy":"(\d+)"/.exec(blob)
    const rid = embedMatch?.[1]
    if (!rid) continue
    const target = rels.get(rid)
    if (!target) continue
    const path = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\.\//, '')}`
    let mediaPath = path.replace(/\\/g, '/')
    if (!mediaPath.startsWith('word/')) mediaPath = `word/${mediaPath}`
    const part = [
      ...doc.package.preservedParts,
      ...Object.values(doc.media).map((m) => ({ path: m.path!, bytes: m.bytes })),
    ].find((p) => {
      const leaf = mediaPath.split('/').pop()!
      return p.path === mediaPath || p.path.endsWith(leaf)
    })
    // EMUs: 914400 per inch; twips: 1440 per inch → twips = emu * 1440 / 914400 = emu / 635
    const cx = Number(cxMatch?.[1] ?? '914400')
    const cy = Number(cyMatch?.[1] ?? '914400')
    const widthTwips = Math.max(1, Math.round(cx / 635))
    const heightTwips = Math.max(1, Math.round(cy / 635))
    const mediaId = nextId('img')
    // bytes filled later in parseDocx when we have zip entries — store placeholder path in media
    doc.media[mediaId] = {
      id: mediaId,
      contentType: 'image/png',
      bytes: part ? part.bytes : new Uint8Array(),
      path: mediaPath,
    }
    return { type: 'image', mediaId, widthTwips, heightTwips }
  }
  return undefined
}

function parseParagraph(
  pNode: Record<string, unknown>,
  doc?: Document,
  rels?: Map<string, string>,
): Paragraph {
  const pPr = findChild(pNode, 'pPr') as Record<string, unknown> | undefined
  const props = parseParagraphProps(pPr)
  const runs: Run[] = []

  for (const [key, value] of Object.entries(pNode)) {
    const name = local(key)
    if (name === 'pPr' || key.startsWith('@_')) continue
    if (name === 'r') {
      for (const r of asArray(value as never)) {
        runs.push(...parseRun(r as Record<string, unknown>, doc, rels))
      }
    } else if (name === 'hyperlink') {
      for (const h of asArray(value as never)) {
        const hNode = h as Record<string, unknown>
        for (const r of findChildren(hNode, 'r')) {
          runs.push(...parseRun(r as Record<string, unknown>, doc, rels))
        }
      }
    }
  }

  if (runs.length === 0) runs.push(createTextRun(''))
  return {
    id: nextId('p'),
    type: 'paragraph',
    props,
    runs,
  }
}

/** Parse a <w:tblBorders>/<w:tcBorders>-style node into edge color strings. */
function parseBordersColors(
  bordersNode: Record<string, unknown> | undefined,
  edges: readonly string[],
): Record<string, string> {
  const out: Record<string, string> = {}
  if (!bordersNode) return out
  for (const edge of edges) {
    const edgeNode = findChild(bordersNode, edge) as Record<string, unknown> | undefined
    if (!edgeNode) continue
    const val = xmlAttrAlt(edgeNode, 'w:val', 'val')
    if (val === 'nil' || val === 'none') continue
    const color = parseColor(xmlAttrAlt(edgeNode, 'w:color', 'color'))
    out[edge] = color ?? '#000000'
  }
  return out
}

function parseTableProps(tblPr: Record<string, unknown> | undefined): Table['props'] {
  const props: Table['props'] = {}
  if (!tblPr) return props
  const tblW = findChild(tblPr, 'tblW') as Record<string, unknown> | undefined
  if (tblW) {
    const type = xmlAttrAlt(tblW, 'w:type', 'type')
    const w = xmlAttrAlt(tblW, 'w:w', 'w')
    if (w && type !== 'pct' && type !== 'auto') props.widthTwips = Number(w)
  }
  const jc = findChild(tblPr, 'jc') as Record<string, unknown> | undefined
  if (jc) {
    const v = xmlAttrAlt(jc, 'w:val', 'val')
    if (v === 'left' || v === 'center' || v === 'right') props.alignment = v
  }
  const tblBorders = findChild(tblPr, 'tblBorders') as Record<string, unknown> | undefined
  const borderColors = parseBordersColors(tblBorders, ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'])
  if (Object.keys(borderColors).length) props.borders = borderColors
  const tblCellSpacing = findChild(tblPr, 'tblCellSpacing') as Record<string, unknown> | undefined
  if (tblCellSpacing) {
    const w = xmlAttrAlt(tblCellSpacing, 'w:w', 'w')
    if (w) props.cellSpacing = Number(w)
  }
  return props
}

function parseCellMargin(tcMar: Record<string, unknown> | undefined): TableCell['props']['margin'] {
  if (!tcMar) return undefined
  const margin: NonNullable<TableCell['props']['margin']> = {}
  for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
    const node = findChild(tcMar, edge) as Record<string, unknown> | undefined
    if (!node) continue
    const w = xmlAttrAlt(node, 'w:w', 'w')
    if (w !== undefined) margin[edge] = Number(w)
  }
  return Object.keys(margin).length > 0 ? margin : undefined
}

function parseTable(tblNode: Record<string, unknown>, doc: Document, rels: Map<string, string>): Table {
  const gridCols: number[] = []
  const tblGrid = findChild(tblNode, 'tblGrid') as Record<string, unknown> | undefined
  if (tblGrid) {
    for (const g of findChildren(tblGrid, 'gridCol')) {
      const w = xmlAttrAlt(g as Record<string, unknown>, 'w:w', 'w')
      if (w) gridCols.push(Number(w))
    }
  }
  const rows: TableRow[] = []
  for (const trNode of findChildren(tblNode, 'tr')) {
    const tr = trNode as Record<string, unknown>
    const trPr = findChild(tr, 'trPr') as Record<string, unknown> | undefined
    const cells: TableCell[] = []
    for (const tcNode of findChildren(tr, 'tc')) {
      const tc = tcNode as Record<string, unknown>
      const tcPr = findChild(tc, 'tcPr') as Record<string, unknown> | undefined
      const gridSpanNode = tcPr
        ? (findChild(tcPr, 'gridSpan') as Record<string, unknown> | undefined)
        : undefined
      const vMergeNode = tcPr ? (findChild(tcPr, 'vMerge') as Record<string, unknown> | undefined) : undefined
      const vAlignNode = tcPr ? (findChild(tcPr, 'vAlign') as Record<string, unknown> | undefined) : undefined
      const tcWNode = tcPr ? (findChild(tcPr, 'tcW') as Record<string, unknown> | undefined) : undefined
      const tcMarNode = tcPr ? (findChild(tcPr, 'tcMar') as Record<string, unknown> | undefined) : undefined
      const tcBordersNode = tcPr ? (findChild(tcPr, 'tcBorders') as Record<string, unknown> | undefined) : undefined
      const shd = tcPr ? (findChild(tcPr, 'shd') as Record<string, unknown> | undefined) : undefined
      const fill = shd ? xmlAttrAlt(shd, 'w:fill', 'fill') : undefined
      const blocks: Paragraph[] = []
      for (const [k, v] of Object.entries(tc)) {
        if (local(k) === 'p') {
          for (const p of asArray(v as never)) blocks.push(parseParagraph(p as Record<string, unknown>, doc, rels))
        }
      }
      if (blocks.length === 0) {
        blocks.push({
          id: nextId('p'),
          type: 'paragraph',
          props: {},
          runs: [createTextRun('')],
        })
      }
      const cell: TableCell = {
        id: nextId('tc'),
        props: {},
        blocks,
      }
      const span = xmlAttrAlt(gridSpanNode, 'w:val', 'val')
      if (span) cell.props.gridSpan = Number(span)
      if (vMergeNode) {
        const vm = xmlAttrAlt(vMergeNode, 'w:val', 'val')
        cell.props.vMerge = vm === 'continue' ? 'continue' : 'restart'
      }
      if (vAlignNode) {
        const va = xmlAttrAlt(vAlignNode, 'w:val', 'val')
        if (va === 'center' || va === 'bottom') cell.props.vAlign = va
        else if (va === 'top') cell.props.vAlign = 'top'
      }
      if (tcWNode) {
        const type = xmlAttrAlt(tcWNode, 'w:type', 'type')
        const w = xmlAttrAlt(tcWNode, 'w:w', 'w')
        if (w && type !== 'pct' && type !== 'auto') cell.props.widthTwips = Number(w)
      }
      const margin = parseCellMargin(tcMarNode)
      if (margin) cell.props.margin = margin
      const cellBorders = parseBordersColors(tcBordersNode, ['top', 'bottom', 'left', 'right'])
      if (Object.keys(cellBorders).length) cell.props.borders = cellBorders
      if (fill && fill !== 'auto') cell.props.shading = `#${fill}`
      cells.push(cell)
    }
    const row: TableRow = { id: nextId('tr'), props: {}, cells }
    if (trPr && findChild(trPr, 'tblHeader') !== undefined) row.props.header = true
    if (trPr) {
      const trHeight = findChild(trPr, 'trHeight') as Record<string, unknown> | undefined
      if (trHeight) {
        const h = xmlAttrAlt(trHeight, 'w:val', 'val')
        if (h) row.props.heightTwips = Number(h)
      }
    }
    rows.push(row)
  }
  const tblPr = findChild(tblNode, 'tblPr') as Record<string, unknown> | undefined
  const table: Table = {
    id: nextId('tbl'),
    type: 'table',
    props: parseTableProps(tblPr),
    rows,
  }
  if (gridCols.length) table.gridCols = gridCols
  return table
}

function parseSectPr(sectPr: Record<string, unknown> | undefined): SectionProperties {
  const properties: SectionProperties = {
    pageSize: { ...DEFAULT_PAGE_SIZE },
    margins: { ...DEFAULT_MARGINS },
    columns: 1,
  }
  if (!sectPr) return properties

  const pgSz = findChild(sectPr, 'pgSz') as Record<string, unknown> | undefined
  if (pgSz) {
    const w = xmlAttrAlt(pgSz, 'w:w', 'w')
    const h = xmlAttrAlt(pgSz, 'w:h', 'h')
    if (w) properties.pageSize.width = Number(w)
    if (h) properties.pageSize.height = Number(h)
  }

  const pgMar = findChild(sectPr, 'pgMar') as Record<string, unknown> | undefined
  if (pgMar) {
    const read = (name: string): number | undefined => {
      const v = xmlAttrAlt(pgMar, `w:${name}`, name)
      return v !== undefined ? Number(v) : undefined
    }
    const top = read('top')
    const right = read('right')
    const bottom = read('bottom')
    const left = read('left')
    const header = read('header')
    const footer = read('footer')
    if (top !== undefined) properties.margins.top = top
    if (right !== undefined) properties.margins.right = right
    if (bottom !== undefined) properties.margins.bottom = bottom
    if (left !== undefined) properties.margins.left = left
    if (header !== undefined) properties.margins.header = header
    if (footer !== undefined) properties.margins.footer = footer
  }

  if (findChild(sectPr, 'titlePg') !== undefined) properties.titlePage = true

  const cols = findChild(sectPr, 'cols') as Record<string, unknown> | undefined
  if (cols) {
    const num = xmlAttrAlt(cols, 'w:num', 'num')
    if (num) properties.columns = Number(num)
  }

  return properties
}

function parseRelationships(relsXml: string | undefined): Map<string, string> {
  const map = new Map<string, string>()
  if (!relsXml) return map
  const root = parseXml(relsXml) as Record<string, unknown>
  const relsRoot = (findChild(root, 'Relationships') ?? root) as Record<string, unknown>
  for (const [k, v] of Object.entries(relsRoot)) {
    if (local(k) !== 'Relationship') continue
    for (const rel of asArray(v as never)) {
      const r = rel as Record<string, unknown>
      const id = xmlAttrAlt(r, 'Id', 'id')
      const target = xmlAttrAlt(r, 'Target', 'target')
      if (id && target) map.set(id, target)
    }
  }
  return map
}

function parseHeaderFooter(xml: string | undefined, doc: Document, rels: Map<string, string>): HeaderFooter | undefined {
  if (!xml) return undefined
  const root = parseXml(xml) as Record<string, unknown>
  const hf = (findChild(root, 'hdr') ?? findChild(root, 'ftr') ?? root) as Record<string, unknown>
  const blocks: Paragraph[] = []
  for (const [k, v] of Object.entries(hf)) {
    if (local(k) === 'p') {
      for (const p of asArray(v as never)) blocks.push(parseParagraph(p as Record<string, unknown>, doc, rels))
    }
  }
  return { id: nextId('hf'), blocks }
}

function parseStyles(stylesXml: string | undefined, doc: Document): void {
  if (!stylesXml) return
  const root = parseXml(stylesXml) as Record<string, unknown>
  const stylesRoot = (findChild(root, 'styles') ?? root) as Record<string, unknown>
  const docDefaults = findChild(stylesRoot, 'docDefaults') as Record<string, unknown> | undefined
  if (docDefaults) {
    const rPrDefault = findChild(docDefaults, 'rPrDefault') as Record<string, unknown> | undefined
    const rPr = rPrDefault
      ? (findChild(rPrDefault, 'rPr') as Record<string, unknown> | undefined)
      : undefined
    doc.styles.docDefaults.character = {
      ...doc.styles.docDefaults.character,
      ...parseRunProps(rPr),
    }
    const pPrDefault = findChild(docDefaults, 'pPrDefault') as Record<string, unknown> | undefined
    const pPr = pPrDefault
      ? (findChild(pPrDefault, 'pPr') as Record<string, unknown> | undefined)
      : undefined
    doc.styles.docDefaults.paragraph = {
      ...doc.styles.docDefaults.paragraph,
      ...parseParagraphProps(pPr),
    }
  }

  for (const styleNode of findChildren(stylesRoot, 'style')) {
    const s = styleNode as Record<string, unknown>
    const type = xmlAttrAlt(s, 'w:type', 'type')
    const styleId = xmlAttrAlt(s, 'w:styleId', 'styleId')
    if (!styleId) continue
    const nameNode = findChild(s, 'name') as Record<string, unknown> | undefined
    const name = (nameNode && (xmlAttrAlt(nameNode, 'w:val', 'val'))) || styleId
    const basedOnNode = findChild(s, 'basedOn') as Record<string, unknown> | undefined
    const basedOn = basedOnNode
      ? xmlAttrAlt(basedOnNode, 'w:val', 'val')
      : undefined
    const pPr = findChild(s, 'pPr') as Record<string, unknown> | undefined
    const rPr = findChild(s, 'rPr') as Record<string, unknown> | undefined

    if (type === 'paragraph' || type === undefined) {
      const style: (typeof doc.styles.paragraphStyles)[string] = {
        id: styleId,
        name,
        paragraph: parseParagraphProps(pPr),
        character: parseRunProps(rPr),
      }
      if (basedOn !== undefined) style.basedOn = basedOn
      doc.styles.paragraphStyles[styleId] = style
    } else if (type === 'character') {
      const style: (typeof doc.styles.characterStyles)[string] = {
        id: styleId,
        name,
        props: parseRunProps(rPr),
      }
      if (basedOn !== undefined) style.basedOn = basedOn
      doc.styles.characterStyles[styleId] = style
    }
  }
}

function collectPreserved(entries: ZipEntries, consumed: Set<string>): Document['package'] {
  const preservedParts = []
  const macroPartPaths: string[] = []
  for (const [path, bytes] of entries) {
    if (consumed.has(path)) continue
    if (isMacroPath(path)) macroPartPaths.push(path)
    preservedParts.push({ path, bytes })
  }
  const contentTypesXml = zipText(entries, '[Content_Types].xml')
  return {
    sourceFormat: 'docx',
    preservedParts,
    macroPartPaths,
    ...(contentTypesXml !== undefined ? { contentTypesXml } : {}),
  }
}

export function parseDocx(bytes: Uint8Array): Document {
  const entries = readZip(bytes)
  const documentXml = zipText(entries, 'word/document.xml')
  assert(documentXml, 'docx_missing', 'word/document.xml is required')

  const doc = createEmptyDocument('docx')
  const stylesXml = zipText(entries, 'word/styles.xml')
  parseStyles(stylesXml, doc)
  doc.numbering = parseNumberingXml(zipText(entries, 'word/numbering.xml'))

  const rels = parseRelationships(zipText(entries, 'word/_rels/document.xml.rels'))

  // Seed media bytes from zip for relationship targets under media/
  for (const [path, data] of entries) {
    if (path.startsWith('word/media/')) {
      const id = nextId('img')
      const lower = path.toLowerCase()
      const contentType = lower.endsWith('.png')
        ? 'image/png'
        : lower.endsWith('.jpg') || lower.endsWith('.jpeg')
          ? 'image/jpeg'
          : lower.endsWith('.gif')
            ? 'image/gif'
            : lower.endsWith('.webp')
              ? 'image/webp'
              : 'application/octet-stream'
      doc.media[id] = { id, contentType, bytes: data, path }
    }
  }

  // Index media by path for drawing resolution
  const mediaByPath = new Map<string, string>()
  for (const [id, item] of Object.entries(doc.media)) {
    if (item.path) mediaByPath.set(item.path, id)
  }

  const root = parseXml(documentXml) as Record<string, unknown>
  const document = findChild(root, 'document') as Record<string, unknown> | undefined
  assert(document, 'docx_missing', 'w:document is required')
  const body = findChild(document, 'body') as Record<string, unknown> | undefined
  assert(body, 'docx_missing', 'w:body is required')

  const blocks: Block[] = []
  for (const [key, value] of Object.entries(body)) {
    const name = local(key)
    if (name === 'p') {
      for (const p of asArray(value as never)) {
        const para = parseParagraph(p as Record<string, unknown>, doc, rels)
        // Remap image mediaIds created during parse to path-based media
        for (const run of para.runs) {
          if (run.content.type !== 'image') continue
          const item = doc.media[run.content.mediaId]
          if (item?.path && mediaByPath.has(item.path)) {
            const realId = mediaByPath.get(item.path)!
            if (realId !== run.content.mediaId) {
              delete doc.media[run.content.mediaId]
              run.content.mediaId = realId
            }
          }
        }
        blocks.push(para)
      }
    } else if (name === 'tbl') {
      for (const t of asArray(value as never)) {
        blocks.push(parseTable(t as Record<string, unknown>, doc, rels))
      }
    }
  }
  if (blocks.length === 0) {
    blocks.push({
      id: nextId('p'),
      type: 'paragraph',
      props: { styleId: 'Normal' },
      runs: [createTextRun('')],
    })
  }

  const section: Section = {
    id: 'sect_1',
    properties: parseSectPr(findChild(body, 'sectPr') as Record<string, unknown> | undefined),
    blocks,
  }

  // Headers / footers via relationships
  for (const [rid, target] of rels) {
    void rid
    const t = target.replace(/^\//, '')
    const path = t.startsWith('word/') ? t : `word/${t}`
    if (path.includes('header')) {
      const xml = zipText(entries, path)
      const hfRels = parseRelationships(zipText(entries, path.replace('word/', 'word/_rels/') + '.rels'))
      const hf = parseHeaderFooter(xml, doc, new Map([...rels, ...hfRels]))
      if (hf && path.includes('header1')) section.header = hf
      else if (hf && !section.header) section.header = hf
    }
    if (path.includes('footer')) {
      const xml = zipText(entries, path)
      const hf = parseHeaderFooter(xml, doc, rels)
      if (hf && path.includes('footer1')) section.footer = hf
      else if (hf && !section.footer) section.footer = hf
    }
  }

  doc.sections = [section]
  doc.package = collectPreserved(entries, new Set(['word/document.xml']))

  return doc
}

export function sniffDocx(bytes: Uint8Array): boolean {
  try {
    const entries = readZip(bytes)
    return entries.has('word/document.xml')
  } catch {
    return false
  }
}
