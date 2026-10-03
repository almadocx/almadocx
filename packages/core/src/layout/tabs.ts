import type { ParagraphProps, TabStop } from '../model/types.js'
import { TWIPS_PER_INCH } from '../util/units.js'

const DEFAULT_TAB_TWIPS = TWIPS_PER_INCH / 2 // 0.5"

/**
 * Resolve the next tab stop position (in twips from left margin / content start)
 * given the current x position in twips.
 */
export function nextTabStopTwips(
  currentXTwips: number,
  tabs: TabStop[] | undefined,
  contentWidthTwips: number,
): {
  position: number
  leader: NonNullable<TabStop['leader']>
  alignment: NonNullable<TabStop['alignment']>
} {
  const stops = [...(tabs ?? [])].sort((a, b) => a.position - b.position)
  for (const stop of stops) {
    if (stop.position > currentXTwips + 1) {
      return {
        position: Math.min(stop.position, contentWidthTwips),
        leader: stop.leader ?? 'none',
        alignment: stop.alignment ?? 'left',
      }
    }
  }
  // Default every 0.5". When past the last stop on the line, return a position
  // beyond contentWidth so layout can wrap and retry on the next line.
  const next = Math.ceil((currentXTwips + 1) / DEFAULT_TAB_TWIPS) * DEFAULT_TAB_TWIPS
  if (next > contentWidthTwips && currentXTwips >= contentWidthTwips - 1) {
    return { position: next, leader: 'none', alignment: 'left' }
  }
  return {
    position: Math.min(next, Math.max(currentXTwips, contentWidthTwips)),
    leader: 'none',
    alignment: 'left',
  }
}

export function defaultTabs(): TabStop[] {
  return Array.from({ length: 12 }, (_, i) => ({
    position: (i + 1) * DEFAULT_TAB_TWIPS,
    alignment: 'left' as const,
    leader: 'none' as const,
  }))
}

export function mergeParagraphTabProps(
  props: ParagraphProps,
  tabs: TabStop[],
): ParagraphProps {
  return { ...props, tabs }
}
