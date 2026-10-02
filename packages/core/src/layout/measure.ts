import { fontStackCss } from '../fonts/substitutes.js'

export interface TextMetrics {
  width: number
  ascent: number
  descent: number
}

export interface TextMeasurer {
  measure(text: string, font: string): TextMetrics
}

/** Approximate measurer for Node/tests when canvas is unavailable. */
export function createApproximateMeasurer(avgCharWidthEm = 0.5): TextMeasurer {
  return {
    measure(text: string, font: string): TextMetrics {
      const sizeMatch = /(\d+(?:\.\d+)?)px/.exec(font)
      const size = sizeMatch ? Number(sizeMatch[1]) : 16
      let units = 0
      for (const ch of text) {
        const code = ch.codePointAt(0)!
        units += code > 0xff ? 1 : avgCharWidthEm
      }
      return {
        width: units * size,
        ascent: size * 0.8,
        descent: size * 0.2,
      }
    },
  }
}

export function fontCss(
  fontFamily: string,
  fontSizePt: number,
  bold?: boolean,
  italic?: boolean,
  dpi = 96,
): string {
  const px = (fontSizePt * dpi) / 72
  const style = italic ? 'italic' : 'normal'
  const weight = bold ? '700' : '400'
  return `${style} ${weight} ${px}px ${fontStackCss(fontFamily)}`
}
