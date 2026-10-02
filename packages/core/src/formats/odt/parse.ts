import {
  createEmptyDocument,
  type Block,
  type CharacterProps,
  type Document,
  type Paragraph,
  type ParagraphProps,
  type Run,
  type Table,
  type TableCell,
  type TableRow,
} from '../../model/types.js'
import { createTextRun } from '../../model/text.js'
import { isMacroPath, readZip, zipText } from '../../io/zip.js'
import { asArray, parseXml, xmlAttr, xmlText } from '../../io/xml.js'
import { parseOdfLengthToTwips } from '../../util/units.js'
import { nextId } from '../../util/id.js'
import { assert } from '../../util/assert.js'

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

type StyleMaps = {
  paragraph: Record<string, ParagraphProps>
  text: Record<string, CharacterProps>
}

function parseTextStyleProps(node: Record<string, unknown>): CharacterProps {
  const props: CharacterProps = {}
  const tp = findChild(node, 'text-properties') as Record<string, unknown> | undefined
  if (!tp) return props
  const weight = xmlAttr(tp, 'fo:font-weight') ?? xmlAttr(tp, 'font-weight')
  if (weight === 'bold' || weight === '700') props.bold = true
  const style = xmlAttr(tp, 'fo:font-style') ?? xmlAttr(tp, 'font-style')
  if (style === 'italic') props.italic = true
  const underline = xmlAttr(tp, 'style:text-underline-style')
  if (underline && underline !== 'none') props.underline = true
  const lineThrough = xmlAttr(tp, 'style:text-line-through-style')
  if (lineThrough && lineThrough !== 'none') props.strike = true
  const size = xmlAttr(tp, 'fo:font-size') ?? xmlAttr(tp, 'font-size')
  if (size) {
    const pt = Number.parseFloat(size)
    if (!Number.isNaN(pt)) props.fontSizePt = pt
  }
  const color = xmlAttr(tp, 'fo:color') ?? xmlAttr(tp, 'color')
  if (color) props.color = color
  const font =
    xmlAttr(tp, 'style:font-name') ??
    xmlAttr(tp, 'fo:font-family') ??
    xmlAttr(tp, 'font-family')
  if (font) props.fontFamily = font.replace(/'/g, '')
  return props
}

function parseParaStyleProps(node: Record<string, unknown>): ParagraphProps {
  const props: ParagraphProps = {}
  const pp = findChild(node, 'paragraph-properties') as Record<string, unknown> | undefined
  if (!pp) return props
  const align = xmlAttr(pp, 'fo:text-align') ?? xmlAttr(pp, 'text-align')
  if (align === 'start' || align === 'left') props.alignment = 'left'
  if (align === 'end' || align === 'right') props.alignment = 'right'
  if (align === 'center') props.alignment = 'center'
  if (align === 'justify') props.alignment = 'justify'
  const marginLeft = xmlAttr(pp, 'fo:margin-left')
  const marginRight = xmlAttr(pp, 'fo:margin-right')
  const marginTop = xmlAttr(pp, 'fo:margin-top')
  const marginBottom = xmlAttr(pp, 'fo:margin-bottom')
  const textIndent = xmlAttr(pp, 'fo:text-indent')
  const ml = parseOdfLengthToTwips(marginLeft)
  const mr = parseOdfLengthToTwips(marginRight)
  const mt = parseOdfLengthToTwips(marginTop)
  const mb = parseOdfLengthToTwips(marginBottom)
  const ti = parseOdfLengthToTwips(textIndent)
  if (ml !== undefined) props.indentLeft = ml
  if (mr !== undefined) props.indentRight = mr
  if (mt !== undefined) props.spacingBefore = mt
  if (mb !== undefined) props.spacingAfter = mb
  if (ti !== undefined) props.indentFirstLine = ti
  const lineHeight = xmlAttr(pp, 'fo:line-height')
  if (lineHeight?.endsWith('%')) {
    props.lineSpacing = Number.parseFloat(lineHeight) / 100
    props.lineSpacingRule = 'auto'
  }
  return props
}

function parseStyles(stylesXml: string | undefined): StyleMaps {
  const maps: StyleMaps = { paragraph: {}, text: {} }
  if (!stylesXml) return maps
  const root = parseXml(stylesXml) as Record<string, unknown>
  const stylesRoot = (findChild(root, 'document-styles') ?? root) as Record<string, unknown>
  const auto = findChild(stylesRoot, 'styles') as Record<string, unknown> | undefined
  const containers = [stylesRoot, auto].filter(Boolean) as Record<string, unknown>[]

  for (const container of containers) {
    for (const styleNode of asArray(findChild(container, 'style') as never)) {
      const s = styleNode as Record<string, unknown>
      const name = xmlAttr(s, 'style:name') ?? xmlAttr(s, 'name')
      const family = xmlAttr(s, 'style:family') ?? xmlAttr(s, 'family')
      if (!name) continue
      if (family === 'paragraph') {
        maps.paragraph[name] = {
          ...parseParaStyleProps(s),
          ...{ styleId: name },
        }
        // also text props on paragraph styles
        maps.text[`p:${name}`] = parseTextStyleProps(s)
      } else if (family === 'text') {
        maps.text[name] = parseTextStyleProps(s)
      }
    }
  }
  return maps
}

function parseInline(
  node: Record<string, unknown>,
  styleMaps: StyleMaps,
  inherited: CharacterProps,
): Run[] {
  const runs: Run[] = []
  const styleName = xmlAttr(node, 'text:style-name') ?? xmlAttr(node, 'style-name')
  const props: CharacterProps = {
    ...inherited,
    ...(styleName ? styleMaps.text[styleName] ?? {} : {}),
  }

  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith('@_')) continue
    const name = local(key)
    if (name === 'span') {
      for (const span of asArray(value as never)) {
        runs.push(...parseInline(span as Record<string, unknown>, styleMaps, props))
      }
    } else if (name === 's') {
      const s = value as Record<string, unknown>
      const c = Number(xmlAttr(s, 'text:c') ?? xmlAttr(s, 'c') ?? '1')
      runs.push(createTextRun(' '.repeat(c), props))
    } else if (name === 'tab') {
      runs.push({ id: nextId('r'), props, content: { type: 'tab' } })
    } else if (name === 'line-break') {
      runs.push({ id: nextId('r'), props, content: { type: 'break', breakType: 'line' } })
    } else if (key === '#text' || name === 'text') {
      const t = xmlText(value)
      if (t) runs.push(createTextRun(t, props))
    }
  }

  // Direct text node on paragraph/heading
  if (typeof node['#text'] === 'string' && runs.length === 0) {
    runs.push(createTextRun(String(node['#text']), props))
  }

  return runs
}

function parseParagraphNode(
  node: Record<string, unknown>,
  styleMaps: StyleMaps,
): Paragraph {
  const styleName = xmlAttr(node, 'text:style-name') ?? xmlAttr(node, 'style-name')
  const props: ParagraphProps = {
    ...(styleName ? styleMaps.paragraph[styleName] ?? { styleId: styleName } : {}),
  }
  const charInherited = styleName ? styleMaps.text[`p:${styleName}`] ?? {} : {}

  const runs: Run[] = []
  // Walk children
  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith('@_')) continue
    const name = local(key)
    if (name === 'span' || name === 'a') {
      for (const child of asArray(value as never)) {
        runs.push(...parseInline(child as Record<string, unknown>, styleMaps, charInherited))
      }
    } else if (name === 's') {
      for (const child of asArray(value as never)) {
        const s = child as Record<string, unknown>
        const c = Number(xmlAttr(s, 'text:c') ?? '1')
        runs.push(createTextRun(' '.repeat(Number.isFinite(c) ? c : 1), charInherited))
      }
    } else if (name === 'tab') {
      runs.push({ id: nextId('r'), props: charInherited, content: { type: 'tab' } })
    } else if (name === 'line-break') {
      runs.push({
        id: nextId('r'),
        props: charInherited,
        content: { type: 'break', breakType: 'line' },
      })
    } else if (key === '#text') {
      runs.push(createTextRun(String(value), charInherited))
    }
  }

  if (runs.length === 0) runs.push(createTextRun(''))
  return { id: nextId('p'), type: 'paragraph', props, runs }
}

function collectOdtBlocks(
  container: Record<string, unknown>,
  styleMaps: StyleMaps,
  blocks: Block[],
  doc: Document,
  listLevel: number,
): void {
  for (const [key, value] of Object.entries(container)) {
    const name = local(key)
    if (name === 'p' || name === 'h') {
      for (const p of asArray(value as never)) {
        blocks.push(parseParagraphNode(p as Record<string, unknown>, styleMaps))
      }
    } else if (name === 'list') {
      for (const list of asArray(value as never)) {
        const listNode = list as Record<string, unknown>
        const styleName = xmlAttr(listNode, 'text:style-name') ?? ''
        const numbered = /num|number|outline/i.test(styleName)
        const numId = numbered ? '2' : '1'
        for (const item of asArray(findChild(listNode, 'list-item') as never)) {
          const itemNode = item as Record<string, unknown>
          // Direct paragraphs in list-item
          for (const [ik, iv] of Object.entries(itemNode)) {
            if (local(ik) === 'p' || local(ik) === 'h') {
              for (const p of asArray(iv as never)) {
                const para = parseParagraphNode(p as Record<string, unknown>, styleMaps)
                para.props.numPr = { numId, ilvl: listLevel }
                blocks.push(para)
              }
            } else if (local(ik) === 'list') {
              collectOdtBlocks({ [ik]: iv }, styleMaps, blocks, doc, listLevel + 1)
            }
          }
        }
      }
    } else if (name === 'table') {
      for (const t of asArray(value as never)) {
        blocks.push(parseOdtTable(t as Record<string, unknown>, styleMaps))
      }
    }
  }
}

function parseOdtTable(tableNode: Record<string, unknown>, styleMaps: StyleMaps): Table {
  const rows: TableRow[] = []
  for (const rowNode of asArray(findChild(tableNode, 'table-row') as never)) {
    const row = rowNode as Record<string, unknown>
    const cells: TableCell[] = []
    for (const [k, v] of Object.entries(row)) {
      if (local(k) !== 'table-cell' && local(k) !== 'covered-table-cell') continue
      for (const cellNode of asArray(v as never)) {
        const cell = cellNode as Record<string, unknown>
        const span = Number(xmlAttr(cell, 'table:number-columns-spanned') ?? '1')
        const blocks: Paragraph[] = []
        for (const [ck, cv] of Object.entries(cell)) {
          if (local(ck) === 'p') {
            for (const p of asArray(cv as never)) {
              blocks.push(parseParagraphNode(p as Record<string, unknown>, styleMaps))
            }
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
        const tc: TableCell = { id: nextId('tc'), props: {}, blocks }
        if (span > 1) tc.props.gridSpan = span
        if (local(k) === 'covered-table-cell') tc.props.vMerge = 'continue'
        cells.push(tc)
      }
    }
    rows.push({ id: nextId('tr'), props: {}, cells })
  }
  return { id: nextId('tbl'), type: 'table', props: {}, rows }
}

export function parseOdt(bytes: Uint8Array): Document {
  const entries = readZip(bytes)
  const contentXml = zipText(entries, 'content.xml')
  assert(contentXml, 'odt_missing', 'content.xml is required')
  const mime = zipText(entries, 'mimetype')
  if (mime) {
    assert(
      mime.includes('application/vnd.oasis.opendocument.text'),
      'odt_mime',
      'not an ODT package',
    )
  }

  const doc = createEmptyDocument('odt')
  const styleMaps = parseStyles(zipText(entries, 'styles.xml'))
  // Also parse automatic styles from content.xml
  const root = parseXml(contentXml) as Record<string, unknown>
  const contentRoot = (findChild(root, 'document-content') ?? root) as Record<string, unknown>
  const autoStyles = findChild(contentRoot, 'automatic-styles') as Record<string, unknown> | undefined
  if (autoStyles) {
    for (const styleNode of asArray(findChild(autoStyles, 'style') as never)) {
      const s = styleNode as Record<string, unknown>
      const name = xmlAttr(s, 'style:name') ?? xmlAttr(s, 'name')
      const family = xmlAttr(s, 'style:family') ?? xmlAttr(s, 'family')
      if (!name) continue
      if (family === 'paragraph') {
        styleMaps.paragraph[name] = { ...parseParaStyleProps(s), styleId: name }
        styleMaps.text[`p:${name}`] = parseTextStyleProps(s)
      } else if (family === 'text') {
        styleMaps.text[name] = parseTextStyleProps(s)
      }
    }
  }

  const body = findChild(contentRoot, 'body') as Record<string, unknown> | undefined
  assert(body, 'odt_missing', 'office:body is required')
  const text = findChild(body, 'text') as Record<string, unknown> | undefined
  assert(text, 'odt_missing', 'office:text is required')

  const blocks: Block[] = []
  collectOdtBlocks(text, styleMaps, blocks, doc, 0)

  if (blocks.length === 0) {
    blocks.push({
      id: nextId('p'),
      type: 'paragraph',
      props: { styleId: 'Standard' },
      runs: [createTextRun('')],
    })
  }

  // Master page header/footer from styles.xml
  const stylesRoot = parseXml(zipText(entries, 'styles.xml') ?? '<r/>') as Record<string, unknown>
  const stylesDoc = (findChild(stylesRoot, 'document-styles') ?? stylesRoot) as Record<string, unknown>
  const masterStyles = findChild(stylesDoc, 'master-styles') as Record<string, unknown> | undefined
  const masterPage = masterStyles
    ? (findChild(masterStyles, 'master-page') as Record<string, unknown> | undefined)
    : undefined
  let header
  let footer
  if (masterPage) {
    const h = findChild(masterPage, 'header') as Record<string, unknown> | undefined
    const f = findChild(masterPage, 'footer') as Record<string, unknown> | undefined
    if (h) {
      const hb: Paragraph[] = []
      for (const [k, v] of Object.entries(h)) {
        if (local(k) === 'p') {
          for (const p of asArray(v as never)) hb.push(parseParagraphNode(p as Record<string, unknown>, styleMaps))
        }
      }
      if (hb.length) header = { id: nextId('hf'), blocks: hb }
    }
    if (f) {
      const fb: Paragraph[] = []
      for (const [k, v] of Object.entries(f)) {
        if (local(k) === 'p') {
          for (const p of asArray(v as never)) fb.push(parseParagraphNode(p as Record<string, unknown>, styleMaps))
        }
      }
      if (fb.length) footer = { id: nextId('hf'), blocks: fb }
    }
  }

  doc.sections = [
    {
      id: 'sect_1',
      properties: doc.sections[0]!.properties,
      blocks,
      ...(header ? { header } : {}),
      ...(footer ? { footer } : {}),
    },
  ]

  const preservedParts = []
  const macroPartPaths: string[] = []
  for (const [path, b] of entries) {
    if (path === 'content.xml') continue
    if (isMacroPath(path)) macroPartPaths.push(path)
    preservedParts.push({ path, bytes: b })
  }
  doc.package = { sourceFormat: 'odt', preservedParts, macroPartPaths }
  return doc
}

export function sniffOdt(bytes: Uint8Array): boolean {
  try {
    const entries = readZip(bytes)
    return entries.has('content.xml') && entries.has('mimetype')
  } catch {
    return false
  }
}
