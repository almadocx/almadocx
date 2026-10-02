/**
 * Lightweight markdown → paragraph fragment for plain-text pastes that look like MD.
 * Handles headings, bullets, numbered lists, bold/italic markers.
 */

import {
  createEmptyParagraph,
  createTextRun,
  type CharacterProps,
  type Paragraph,
} from '@almadocx/core'
import type { PastedFragment } from './sanitize.js'

const HEADING_RE = /^(#{1,6})\s+(.*)$/
const UL_RE = /^(\s*)([-*+])\s+(.*)$/
const OL_RE = /^(\s*)(\d+)[.)]\s+(.*)$/

function looksLikeMarkdown(text: string): boolean {
  const lines = text.split(/\r?\n/)
  let signals = 0
  for (const line of lines.slice(0, 40)) {
    if (HEADING_RE.test(line) || UL_RE.test(line) || OL_RE.test(line)) signals += 1
    if (/\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_/.test(line)) signals += 1
  }
  return signals >= 1
}

function parseInline(text: string, base: CharacterProps = {}): Paragraph['runs'] {
  const runs: Paragraph['runs'] = []
  const re = /(\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_|`[^`]+`)/g
  let last = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(text))) {
    if (m.index > last) {
      runs.push(createTextRun(text.slice(last, m.index), base))
    }
    const token = m[0]!
    if (token.startsWith('**') || token.startsWith('__')) {
      runs.push(createTextRun(token.slice(2, -2), { ...base, bold: true }))
    } else if (token.startsWith('*') || token.startsWith('_')) {
      runs.push(createTextRun(token.slice(1, -1), { ...base, italic: true }))
    } else if (token.startsWith('`')) {
      runs.push(
        createTextRun(token.slice(1, -1), {
          ...base,
          fontFamily: 'Consolas',
          fontSizePt: (base.fontSizePt ?? 11) * 0.95,
        }),
      )
    }
    last = m.index + token.length
  }
  if (last < text.length) runs.push(createTextRun(text.slice(last), base))
  if (runs.length === 0) runs.push(createTextRun(''))
  return runs
}

export function markdownToFragment(text: string): PastedFragment | undefined {
  if (!looksLikeMarkdown(text)) return undefined
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n')
  const paragraphs: Paragraph[] = []

  for (const line of lines) {
    const h = HEADING_RE.exec(line)
    if (h) {
      const level = h[1]!.length
      const p = createEmptyParagraph(`Heading${String(level)}`)
      const size = 24 - (level - 1) * 2
      p.runs = parseInline(h[2] ?? '', { bold: true, fontSizePt: size })
      paragraphs.push(p)
      continue
    }
    const ul = UL_RE.exec(line)
    if (ul) {
      const p = createEmptyParagraph()
      p.props = { numPr: { numId: '1', ilvl: 0 } }
      p.runs = parseInline(ul[3] ?? '')
      paragraphs.push(p)
      continue
    }
    const ol = OL_RE.exec(line)
    if (ol) {
      const p = createEmptyParagraph()
      p.props = { numPr: { numId: '2', ilvl: 0 } }
      p.runs = parseInline(ol[3] ?? '')
      paragraphs.push(p)
      continue
    }
    const p = createEmptyParagraph()
    p.runs = parseInline(line)
    paragraphs.push(p)
  }

  if (paragraphs.length === 0) paragraphs.push(createEmptyParagraph())
  return { paragraphs }
}
