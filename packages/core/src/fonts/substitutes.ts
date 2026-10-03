/**
 * Metric-compatible font substitutes for common Office faces.
 * Layout and paint must resolve through the same mapping.
 */
export const FONT_SUBSTITUTES: Record<string, string> = {
  // Prefer fonts commonly installed on Linux; Carlito/Caladea when present.
  Calibri: 'Carlito',
  Cambria: 'Caladea',
  'Times New Roman': 'Liberation Serif',
  Arial: 'Liberation Sans',
  'Courier New': 'Liberation Mono',
  Helvetica: 'Liberation Sans',
  Georgia: 'Liberation Serif',
  Verdana: 'DejaVu Sans',
  Tahoma: 'DejaVu Sans',
  Symbol: 'DejaVu Sans',
  Wingdings: 'DejaVu Sans',
}

const SANS_FACES = new Set([
  'Carlito',
  'Liberation Sans',
  'DejaVu Sans',
  'Arial',
  'Helvetica',
  'Calibri',
  'Verdana',
  'Tahoma',
])

export const FONT_FALLBACK_STACK_SERIF = [
  'Liberation Serif',
  'DejaVu Serif',
  'Times New Roman',
  'serif',
] as const

export const FONT_FALLBACK_STACK_SANS = [
  'Carlito',
  'Liberation Sans',
  'DejaVu Sans',
  'Arial',
  'sans-serif',
] as const

/** @deprecated use FONT_FALLBACK_STACK_SERIF / _SANS */
export const FONT_FALLBACK_STACK = FONT_FALLBACK_STACK_SERIF

export function resolveFontFamily(requested: string | undefined): string {
  const name = (requested ?? 'Liberation Serif').replace(/['"]/g, '').trim()
  const substituted = FONT_SUBSTITUTES[name] ?? name
  return substituted
}

export function fontStackCss(requested: string | undefined): string {
  const primary = resolveFontFamily(requested)
  const raw = (requested ?? '').replace(/['"]/g, '').trim()
  const sans = SANS_FACES.has(primary) || SANS_FACES.has(raw)
  const stack = sans ? FONT_FALLBACK_STACK_SANS : FONT_FALLBACK_STACK_SERIF
  const rest = stack.filter((f) => f !== primary)
  return [primary, ...rest].map((f) => (f.includes(' ') ? `"${f}"` : f)).join(', ')
}
