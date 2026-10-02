/** Twips: 1/20 of a point. 1440 twips = 1 inch. */
export const TWIPS_PER_INCH = 1440
export const TWIPS_PER_POINT = 20

export function inchesToTwips(inches: number): number {
  return Math.round(inches * TWIPS_PER_INCH)
}

export function pointsToTwips(points: number): number {
  return Math.round(points * TWIPS_PER_POINT)
}

export function twipsToPx(twips: number, dpi = 96): number {
  return (twips / TWIPS_PER_INCH) * dpi
}

export function pxToTwips(px: number, dpi = 96): number {
  return Math.round((px / dpi) * TWIPS_PER_INCH)
}

/** ODF uses 1/100mm (hm). */
export function hmToTwips(hm: number): number {
  // 1 inch = 25.4 mm = 2540 hm; 1440 twips/inch
  return Math.round((hm / 2540) * TWIPS_PER_INCH)
}

export function twipsToHm(twips: number): number {
  return Math.round((twips / TWIPS_PER_INCH) * 2540)
}

export function parseOdfLengthToTwips(value: string | undefined): number | undefined {
  if (value === undefined || value === '') return undefined
  const m = /^(-?\d+(?:\.\d+)?)([a-z]+)?$/i.exec(value.trim())
  if (!m) return undefined
  const n = Number(m[1])
  const unit = (m[2] ?? 'pt').toLowerCase()
  switch (unit) {
    case 'cm':
      return Math.round((n / 2.54) * TWIPS_PER_INCH)
    case 'mm':
      return Math.round((n / 25.4) * TWIPS_PER_INCH)
    case 'in':
      return Math.round(n * TWIPS_PER_INCH)
    case 'pt':
      return pointsToTwips(n)
    case 'pc':
      return pointsToTwips(n * 12)
    case 'px':
      return pxToTwips(n)
    default:
      return undefined
  }
}
