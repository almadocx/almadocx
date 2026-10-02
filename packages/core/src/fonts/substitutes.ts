/**
 * Metric-compatible font substitutes for common Office faces.
 * Layout and paint must resolve through the same mapping.
 */
export const FONT_SUBSTITUTES: Record<string, string> = {
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

export const FONT_FALLBACK_STACK = [
  'Liberation Serif',
  'Liberation Sans',
  'DejaVu Serif',
  'DejaVu Sans',
  'serif',
] as const

export function resolveFontFamily(requested: string | undefined): string {
  const name = (requested ?? 'Liberation Serif').replace(/['"]/g, '').trim()
  const substituted = FONT_SUBSTITUTES[name] ?? name
  return substituted
}

export function fontStackCss(requested: string | undefined): string {
  const primary = resolveFontFamily(requested)
  const rest = FONT_FALLBACK_STACK.filter((f) => f !== primary)
  return [primary, ...rest].map((f) => (f.includes(' ') ? `"${f}"` : f)).join(', ')
}
