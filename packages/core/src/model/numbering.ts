import type {
  Document,
  NumberFormat,
  NumberingDefinitions,
  NumberingLevel,
  Paragraph,
} from './types.js'

export interface ResolvedListMarker {
  text: string
  indentLeft: number
  hanging: number
  fontFamily?: string
}

function toRoman(n: number, upper: boolean): string {
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
  const syms = upper
    ? ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
    : ['m', 'cm', 'd', 'cd', 'c', 'xc', 'l', 'xl', 'x', 'ix', 'v', 'iv', 'i']
  let x = Math.max(1, Math.floor(n))
  let out = ''
  for (let i = 0; i < vals.length; i++) {
    while (x >= vals[i]!) {
      out += syms[i]!
      x -= vals[i]!
    }
  }
  return out
}

function toAlpha(n: number, upper: boolean): string {
  let x = Math.max(1, Math.floor(n))
  let out = ''
  while (x > 0) {
    x -= 1
    out = String.fromCharCode((upper ? 65 : 97) + (x % 26)) + out
    x = Math.floor(x / 26)
  }
  return out
}

export function formatNumber(format: NumberFormat, value: number): string {
  switch (format) {
    case 'decimal':
      return String(value)
    case 'lowerLetter':
      return toAlpha(value, false)
    case 'upperLetter':
      return toAlpha(value, true)
    case 'lowerRoman':
      return toRoman(value, false)
    case 'upperRoman':
      return toRoman(value, true)
    case 'bullet':
      return '•'
    case 'none':
      return ''
  }
}

export function renderLevelText(levelText: string, counters: number[], levels: NumberingLevel[]): string {
  let out = levelText
  for (let i = 0; i < 9; i++) {
    const token = `%${String(i + 1)}`
    if (!out.includes(token)) continue
    const level = levels.find((l) => l.ilvl === i)
    const n = counters[i] ?? level?.start ?? 1
    const formatted = level ? formatNumber(level.format, n) : String(n)
    out = out.split(token).join(formatted)
  }
  return out
}

/**
 * Walk section blocks and compute list marker text for each numbered paragraph.
 * Counters reset when numId changes or when a higher level restarts a lower one.
 */
export function resolveListMarkers(
  doc: Document,
  sectionIndex: number,
): Map<number, ResolvedListMarker> {
  const section = doc.sections[sectionIndex]
  const out = new Map<number, ResolvedListMarker>()
  if (!section) return out

  const defs = doc.numbering
  /** numId → counters by ilvl */
  const counterByNum = new Map<string, number[]>()

  for (let bi = 0; bi < section.blocks.length; bi++) {
    const block = section.blocks[bi]
    if (!block || block.type !== 'paragraph' || !block.props.numPr) continue
    const { numId, ilvl } = block.props.numPr
    const instance = defs.nums[numId]
    if (!instance) continue
    const abs = defs.abstractNums[instance.abstractNumId]
    if (!abs) continue
    const level = abs.levels.find((l) => l.ilvl === ilvl) ?? abs.levels[0]
    if (!level) continue

    let counters = counterByNum.get(numId)
    if (!counters) {
      counters = abs.levels.map((l) => {
        const override = instance.startOverrides?.[l.ilvl]
        return override ?? l.start
      })
      // Initialize all to start-1 so first increment yields start
      counters = counters.map((c) => c - 1)
      counterByNum.set(numId, counters)
    }

    // Increment this level; reset deeper levels
    counters[ilvl] = (counters[ilvl] ?? (level.start - 1)) + 1
    for (let d = ilvl + 1; d < counters.length; d++) {
      const deeper = abs.levels.find((l) => l.ilvl === d)
      counters[d] = (deeper?.start ?? 1) - 1
    }

    const text =
      level.format === 'bullet' && !level.levelText.includes('%')
        ? level.levelText || '•'
        : renderLevelText(level.levelText, counters, abs.levels)

    const marker: ResolvedListMarker = {
      text: text + (level.format === 'bullet' ? '' : '\u00a0'),
      indentLeft: level.indentLeft ?? 720,
      hanging: level.hanging ?? 360,
    }
    if (level.fontFamily) marker.fontFamily = level.fontFamily
    out.set(bi, marker)
  }

  return out
}

export function getParagraphListMarker(
  doc: Document,
  paragraph: Paragraph,
  blockIndex: number,
  sectionIndex: number,
): ResolvedListMarker | undefined {
  if (!paragraph.props.numPr) return undefined
  return resolveListMarkers(doc, sectionIndex).get(blockIndex)
}

export function lookupLevel(
  defs: NumberingDefinitions,
  numId: string,
  ilvl: number,
): NumberingLevel | undefined {
  const instance = defs.nums[numId]
  if (!instance) return undefined
  const abs = defs.abstractNums[instance.abstractNumId]
  return abs?.levels.find((l) => l.ilvl === ilvl)
}
