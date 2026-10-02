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
): { position: number; leader: NonNullable<TabStop['leader']> } {
  const stops = [...(tabs ?? [])].sort((a, b) => a.position - b.position)
  for (const stop of stops) {
    if (stop.position > currentXTwips + 1) {
      return {
        position: Math.min(stop.position, contentWidthTwips),
        leader: stop.leader ?? 'none',
      }
    }
  }
  // Default every 0.5"
  const next = Math.ceil((currentXTwips + 1) / DEFAULT_TAB_TWIPS) * DEFAULT_TAB_TWIPS
  return {
    position: Math.min(next, Math.max(currentXTwips, contentWidthTwips)),
    leader: 'none',
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
