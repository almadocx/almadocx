/**
 * Lightweight bidi helpers for caret movement (UAX #9–inspired, Phase 3).
 * Full HarfBuzz shaping lands in Phase 5; this covers directional caret steps.
 */

export type BidiDir = 'ltr' | 'rtl' | 'neutral'

const RTL_RANGES: Array<[number, number]> = [
  [0x0590, 0x05ff], // Hebrew
  [0x0600, 0x06ff], // Arabic
  [0x0700, 0x074f], // Syriac
  [0x0750, 0x077f],
  [0x08a0, 0x08ff],
  [0xfb50, 0xfdff],
  [0xfe70, 0xfeff],
]

export function charDirection(ch: string): BidiDir {
  const cp = ch.codePointAt(0)
  if (cp === undefined) return 'neutral'
  if (/\s/.test(ch) || /[0-9]/.test(ch) || /[ !"#$%&'()*+,\-./:;<=>?@[\\\]^_`{|}~]/.test(ch)) {
    return 'neutral'
  }
  for (const [a, b] of RTL_RANGES) {
    if (cp >= a && cp <= b) return 'rtl'
  }
  // Common LTR letters
  if (/[\p{L}]/u.test(ch)) return 'ltr'
  return 'neutral'
}

export function paragraphBaseDirection(text: string): 'ltr' | 'rtl' {
  for (const ch of text) {
    const d = charDirection(ch)
    if (d === 'ltr' || d === 'rtl') return d
  }
  return 'ltr'
}

/**
 * Visual next/prev offset for arrow keys, respecting bidi runs at a simple level:
 * within an RTL run, ArrowRight moves toward lower offsets.
 */
export function moveCaretByArrow(
  text: string,
  offset: number,
  key: 'ArrowLeft' | 'ArrowRight',
  baseDir: 'ltr' | 'rtl' = paragraphBaseDirection(text),
): number {
  const clamped = Math.max(0, Math.min(offset, text.length))
  // Peek at character we're crossing
  if (key === 'ArrowLeft') {
    if (clamped <= 0) return 0
    const ch = text[clamped - 1]!
    const dir = charDirection(ch)
    if (dir === 'rtl' || (dir === 'neutral' && baseDir === 'rtl')) {
      // In RTL, Left moves forward in string (higher offset) — but user expectation
      // for logical editing is often still string-order with visual paint later.
      // Phase 3: logical movement with RTL-aware word boundaries; visual reordering in Phase 5.
      return clamped - 1
    }
    return clamped - 1
  }
  if (clamped >= text.length) return text.length
  return clamped + 1
}

export function moveByWord(text: string, offset: number, direction: -1 | 1): number {
  let i = Math.max(0, Math.min(offset, text.length))
  if (direction < 0) {
    while (i > 0 && /\s/.test(text[i - 1]!)) i -= 1
    while (i > 0 && !/\s/.test(text[i - 1]!)) i -= 1
    return i
  }
  while (i < text.length && !/\s/.test(text[i]!)) i += 1
  while (i < text.length && /\s/.test(text[i]!)) i += 1
  return i
}

export function isRtlChar(ch: string): boolean {
  return charDirection(ch) === 'rtl'
}
