import type { Document, Paragraph } from '../model/types.js'
import {
  clampPosition,
  paragraphPlainText,
  type DocPosition,
  type DocRange,
} from '../model/position.js'
import { applyOp } from '../ops/apply.js'

export interface FindOptions {
  query: string
  caseSensitive?: boolean
  wholeWord?: boolean
}

export interface FindMatch {
  range: DocRange
}

function isWordChar(ch: string): boolean {
  return /[\p{L}\p{N}_]/u.test(ch)
}

function wholeWordAt(text: string, start: number, end: number): boolean {
  const before = start === 0 ? '' : text[start - 1] ?? ''
  const after = end >= text.length ? '' : text[end] ?? ''
  const leftOk = !before || !isWordChar(before)
  const rightOk = !after || !isWordChar(after)
  return leftOk && rightOk
}

/**
 * Find all matches of query across paragraph text in document order.
 */
export function findAll(doc: Document, options: FindOptions): FindMatch[] {
  const q = options.query
  if (!q) return []
  const needle = options.caseSensitive ? q : q.toLocaleLowerCase()
  const matches: FindMatch[] = []

  for (let si = 0; si < doc.sections.length; si++) {
    const section = doc.sections[si]!
    for (let bi = 0; bi < section.blocks.length; bi++) {
      const block = section.blocks[bi]
      if (!block || block.type !== 'paragraph') continue
      const plain = paragraphPlainText(block)
      const hay = options.caseSensitive ? plain : plain.toLocaleLowerCase()
      let from = 0
      while (from <= hay.length - needle.length) {
        const idx = hay.indexOf(needle, from)
        if (idx < 0) break
        const end = idx + needle.length
        if (!options.wholeWord || wholeWordAt(plain, idx, end)) {
          matches.push({
            range: {
              anchor: { sectionIndex: si, blockIndex: bi, offset: idx },
              focus: { sectionIndex: si, blockIndex: bi, offset: end },
            },
          })
        }
        from = idx + Math.max(1, needle.length)
      }
    }
  }
  return matches
}

export function findNext(
  doc: Document,
  options: FindOptions,
  after: DocPosition,
): FindMatch | undefined {
  const all = findAll(doc, options)
  for (const m of all) {
    const start = m.range.anchor
    if (
      start.sectionIndex > after.sectionIndex ||
      (start.sectionIndex === after.sectionIndex && start.blockIndex > after.blockIndex) ||
      (start.sectionIndex === after.sectionIndex &&
        start.blockIndex === after.blockIndex &&
        start.offset >= after.offset)
    ) {
      return m
    }
  }
  return all[0]
}

/**
 * Replace a single match range with text (same paragraph only).
 */
export function replaceMatch(
  doc: Document,
  match: FindMatch,
  replacement: string,
): { doc: Document; caret: DocPosition } {
  const start = clampPosition(doc, match.range.anchor)
  const end = clampPosition(doc, match.range.focus)
  let next = applyOp(doc, {
    type: 'deleteRange',
    range: { anchor: start, focus: end },
  }).doc
  if (replacement.length > 0) {
    next = applyOp(next, {
      type: 'insertText',
      position: start,
      text: replacement,
    }).doc
  }
  return {
    doc: next,
    caret: { ...start, offset: start.offset + replacement.length },
  }
}

/**
 * Replace all matches. Returns count replaced.
 */
export function replaceAll(
  doc: Document,
  options: FindOptions,
  replacement: string,
): { doc: Document; count: number } {
  const matches = findAll(doc, options)
  // Apply from end to start so offsets stay valid within each paragraph;
  // cross-paragraph we still go reverse document order.
  const ordered = [...matches].reverse()
  let current = doc
  let count = 0
  for (const m of ordered) {
    current = replaceMatch(current, m, replacement).doc
    count += 1
  }
  return { doc: current, count }
}

export function extractPlainRange(doc: Document, range: DocRange): string {
  const start = clampPosition(doc, range.anchor)
  const end = clampPosition(doc, range.focus)
  if (start.sectionIndex !== end.sectionIndex) {
    // Multi-section: join with newlines between paragraphs
    const parts: string[] = []
    for (let si = start.sectionIndex; si <= end.sectionIndex; si++) {
      const section = doc.sections[si]
      if (!section) continue
      const biStart = si === start.sectionIndex ? start.blockIndex : 0
      const biEnd = si === end.sectionIndex ? end.blockIndex : section.blocks.length - 1
      for (let bi = biStart; bi <= biEnd; bi++) {
        const block = section.blocks[bi]
        if (!block || block.type !== 'paragraph') continue
        const plain = paragraphPlainText(block)
        const a = si === start.sectionIndex && bi === start.blockIndex ? start.offset : 0
        const b = si === end.sectionIndex && bi === end.blockIndex ? end.offset : plain.length
        parts.push(plain.slice(a, b))
      }
    }
    return parts.join('\n')
  }
  if (start.blockIndex === end.blockIndex) {
    const block = doc.sections[start.sectionIndex]?.blocks[start.blockIndex]
    if (!block || block.type !== 'paragraph') return ''
    return paragraphPlainText(block).slice(start.offset, end.offset)
  }
  const section = doc.sections[start.sectionIndex]
  if (!section) return ''
  const parts: string[] = []
  for (let bi = start.blockIndex; bi <= end.blockIndex; bi++) {
    const block = section.blocks[bi]
    if (!block || block.type !== 'paragraph') continue
    const plain = paragraphPlainText(block)
    const a = bi === start.blockIndex ? start.offset : 0
    const b = bi === end.blockIndex ? end.offset : plain.length
    parts.push(plain.slice(a, b))
  }
  return parts.join('\n')
}

export function paragraphAt(doc: Document, pos: DocPosition): Paragraph | undefined {
  const block = doc.sections[pos.sectionIndex]?.blocks[pos.blockIndex]
  return block?.type === 'paragraph' ? block : undefined
}
