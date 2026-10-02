import type {
  AbstractNumbering,
  NumberFormat,
  NumberingDefinitions,
  NumberingInstance,
  NumberingLevel,
} from '../../model/types.js'
import { emptyNumbering, ensureDefaultNumbering } from '../../model/types.js'
import { asArray, parseXml, xmlAttrAlt, xmlAttrAltOr } from '../../io/xml.js'

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
  return asArray(findChild(node, localName) as never)
}

function mapNumFmt(val: string | undefined): NumberFormat {
  switch (val) {
    case 'decimal':
      return 'decimal'
    case 'lowerLetter':
      return 'lowerLetter'
    case 'upperLetter':
      return 'upperLetter'
    case 'lowerRoman':
      return 'lowerRoman'
    case 'upperRoman':
      return 'upperRoman'
    case 'bullet':
      return 'bullet'
    case 'none':
      return 'none'
    default:
      return val ? 'decimal' : 'bullet'
  }
}

export function parseNumberingXml(xml: string | undefined): NumberingDefinitions {
  if (!xml) return ensureDefaultNumbering(emptyNumbering())
  const root = parseXml(xml) as Record<string, unknown>
  const numbering = (findChild(root, 'numbering') ?? root) as Record<string, unknown>
  const abstractNums: Record<string, AbstractNumbering> = {}
  const nums: Record<string, NumberingInstance> = {}

  for (const node of findChildren(numbering, 'abstractNum')) {
    const a = node as Record<string, unknown>
    const id = xmlAttrAlt(a, 'w:abstractNumId', 'abstractNumId')
    if (!id) continue
    const levels: NumberingLevel[] = []
    for (const lvlNode of findChildren(a, 'lvl')) {
      const lvl = lvlNode as Record<string, unknown>
      const ilvl = Number(xmlAttrAltOr(lvl, '0', 'w:ilvl', 'ilvl'))
      const startNode = findChild(lvl, 'start') as Record<string, unknown> | undefined
      const numFmtNode = findChild(lvl, 'numFmt') as Record<string, unknown> | undefined
      const lvlTextNode = findChild(lvl, 'lvlText') as Record<string, unknown> | undefined
      const pPr = findChild(lvl, 'pPr') as Record<string, unknown> | undefined
      const ind = pPr ? (findChild(pPr, 'ind') as Record<string, unknown> | undefined) : undefined
      const left = ind ? Number(xmlAttrAltOr(ind, '720', 'w:left', 'left')) : 720
      const hanging = ind ? Number(xmlAttrAltOr(ind, '360', 'w:hanging', 'hanging')) : 360
      const rPr = findChild(lvl, 'rPr') as Record<string, unknown> | undefined
      const rFonts = rPr ? (findChild(rPr, 'rFonts') as Record<string, unknown> | undefined) : undefined
      const fontFamily = rFonts
        ? xmlAttrAlt(rFonts, 'w:ascii', 'ascii')
        : undefined
      const level: NumberingLevel = {
        ilvl,
        format: mapNumFmt(xmlAttrAlt(numFmtNode, 'w:val', 'val')),
        levelText: xmlAttrAltOr(lvlTextNode, '%1.', 'w:val', 'val'),
        start: Number(xmlAttrAltOr(startNode, '1', 'w:val', 'val')),
        indentLeft: left,
        hanging,
      }
      if (fontFamily) level.fontFamily = fontFamily
      levels.push(level)
    }
    levels.sort((a, b) => a.ilvl - b.ilvl)
    abstractNums[id] = { id, levels }
  }

  for (const node of findChildren(numbering, 'num')) {
    const n = node as Record<string, unknown>
    const numId = xmlAttrAlt(n, 'w:numId', 'numId')
    const absNode = findChild(n, 'abstractNumId') as Record<string, unknown> | undefined
    const abstractNumId = xmlAttrAlt(absNode, 'w:val', 'val')
    if (!numId || !abstractNumId) continue
    const startOverrides: Record<number, number> = {}
    for (const ov of findChildren(n, 'lvlOverride')) {
      const o = ov as Record<string, unknown>
      const ilvl = Number(xmlAttrAltOr(o, '0', 'w:ilvl', 'ilvl'))
      const startOv = findChild(o, 'startOverride') as Record<string, unknown> | undefined
      const val = xmlAttrAlt(startOv, 'w:val', 'val')
      if (val) startOverrides[ilvl] = Number(val)
    }
    const inst: NumberingInstance = { numId, abstractNumId }
    if (Object.keys(startOverrides).length) inst.startOverrides = startOverrides
    nums[numId] = inst
  }

  return ensureDefaultNumbering({ abstractNums, nums })
}

export function serializeNumberingXml(defs: NumberingDefinitions): string {
  const abs = Object.values(defs.abstractNums)
    .map((a) => {
      const levels = a.levels
        .map((l) => {
          const ind = `<w:pPr><w:ind w:left="${String(l.indentLeft ?? 720)}" w:hanging="${String(l.hanging ?? 360)}"/></w:pPr>`
          return (
            `<w:lvl w:ilvl="${String(l.ilvl)}">` +
            `<w:start w:val="${String(l.start)}"/>` +
            `<w:numFmt w:val="${l.format === 'bullet' ? 'bullet' : l.format}"/>` +
            `<w:lvlText w:val="${escapeXml(l.levelText)}"/>` +
            ind +
            `</w:lvl>`
          )
        })
        .join('')
      return `<w:abstractNum w:abstractNumId="${escapeXml(a.id)}">${levels}</w:abstractNum>`
    })
    .join('')

  const nums = Object.values(defs.nums)
    .map((n) => {
      let overrides = ''
      if (n.startOverrides) {
        for (const [ilvl, start] of Object.entries(n.startOverrides)) {
          overrides += `<w:lvlOverride w:ilvl="${ilvl}"><w:startOverride w:val="${String(start)}"/></w:lvlOverride>`
        }
      }
      return (
        `<w:num w:numId="${escapeXml(n.numId)}">` +
        `<w:abstractNumId w:val="${escapeXml(n.abstractNumId)}"/>` +
        overrides +
        `</w:num>`
      )
    })
    .join('')

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    abs +
    nums +
    `</w:numbering>`
  )
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}
