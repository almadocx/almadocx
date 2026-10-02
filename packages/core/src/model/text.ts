import { nextId } from '../util/id.js'
import type { CharacterProps, Paragraph, Run } from './types.js'
import { hitTestRun, inlineToChar, paragraphPlainText } from './position.js'

export function createTextRun(text: string, props: CharacterProps = {}): Run {
  return {
    id: nextId('r'),
    props: { ...props },
    content: { type: 'text', text },
  }
}

export function createEmptyParagraph(styleId = 'Normal'): Paragraph {
  return {
    id: nextId('p'),
    type: 'paragraph',
    props: { styleId },
    runs: [createTextRun('')],
  }
}

export function splitRun(run: Run, offsetInRun: number): [Run, Run] {
  if (run.content.type !== 'text') {
    if (offsetInRun <= 0) {
      return [createTextRun('', { ...run.props }), { ...run, id: nextId('r') }]
    }
    return [{ ...run, id: nextId('r') }, createTextRun('', { ...run.props })]
  }
  const leftText = run.content.text.slice(0, offsetInRun)
  const rightText = run.content.text.slice(offsetInRun)
  return [
    { ...run, id: nextId('r'), content: { type: 'text', text: leftText } },
    {
      id: nextId('r'),
      props: { ...run.props },
      ...(run.styleId !== undefined ? { styleId: run.styleId } : {}),
      content: { type: 'text', text: rightText },
    },
  ]
}

function sameCharProps(a: CharacterProps, b: CharacterProps): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.fontFamily === b.fontFamily &&
    a.fontSizePt === b.fontSizePt &&
    a.color === b.color &&
    a.highlight === b.highlight &&
    a.verticalAlign === b.verticalAlign
  )
}

export function coalesceRuns(runs: Run[]): Run[] {
  const out: Run[] = []
  for (const run of runs) {
    if (run.content.type === 'text' && run.content.text.length === 0) {
      continue
    }
    const prev = out[out.length - 1]
    if (
      prev &&
      prev.content.type === 'text' &&
      run.content.type === 'text' &&
      sameCharProps(prev.props, run.props) &&
      prev.styleId === run.styleId
    ) {
      out[out.length - 1] = {
        ...prev,
        content: { type: 'text', text: prev.content.text + run.content.text },
      }
      continue
    }
    out.push(run)
  }
  if (out.length === 0) out.push(createTextRun(''))
  return out
}

export function insertTextInParagraph(
  paragraph: Paragraph,
  offset: number,
  text: string,
  props?: CharacterProps,
): Paragraph {
  if (text.length === 0) return paragraph
  const hit = hitTestRun(paragraph, offset)
  const runs = [...paragraph.runs]

  if (hit.run.content.type === 'text') {
    const [left, right] = splitRun(hit.run, hit.offsetInRun)
    const inserted = createTextRun(text, props ?? { ...hit.run.props })
    runs.splice(hit.runIndex, 1, left, inserted, right)
  } else if (hit.offsetInRun === 0) {
    runs.splice(hit.runIndex, 0, createTextRun(text, props ?? {}))
  } else {
    runs.splice(hit.runIndex + 1, 0, createTextRun(text, props ?? {}))
  }

  return { ...paragraph, runs: coalesceRuns(runs) }
}

export function deleteRangeInParagraph(paragraph: Paragraph, start: number, end: number): Paragraph {
  if (end <= start) return paragraph
  const plain = paragraphPlainText(paragraph)
  const kept = plain.slice(0, start) + plain.slice(end)
  const hit = hitTestRun(paragraph, Math.min(start, paragraphPlainText(paragraph).length))
  const run: Run = {
    id: nextId('r'),
    props: { ...hit.run.props },
    ...(hit.run.styleId !== undefined ? { styleId: hit.run.styleId } : {}),
    content: { type: 'text', text: kept },
  }
  return { ...paragraph, runs: coalesceRuns([run]) }
}

/**
 * Apply character marks to [start, end) by splitting runs at boundaries.
 */
export function setMarkInParagraph(
  paragraph: Paragraph,
  start: number,
  end: number,
  mark: Partial<CharacterProps>,
): Paragraph {
  if (end <= start) return paragraph

  type Piece =
    | { kind: 'text'; text: string; props: CharacterProps; styleId?: string }
    | { kind: 'atomic'; run: Run }

  const pieces: Piece[] = []
  let cursor = 0

  for (const run of paragraph.runs) {
    if (run.content.type === 'image') {
      const runStart = cursor
      const runEnd = cursor + 1
      if (runStart >= start && runEnd <= end) {
        pieces.push({
          kind: 'atomic',
          run: { ...run, props: { ...run.props, ...mark } },
        })
      } else {
        pieces.push({ kind: 'atomic', run })
      }
      cursor = runEnd
      continue
    }

    const chunk = inlineToChar(run)
    if (run.content.type !== 'text') {
      const runStart = cursor
      const runEnd = cursor + chunk.length
      const inside = runStart >= start && runEnd <= end
      pieces.push({
        kind: 'atomic',
        run: {
          ...run,
          props: inside ? { ...run.props, ...mark } : { ...run.props },
        },
      })
      cursor = runEnd
      continue
    }

    let i = 0
    while (i < chunk.length) {
      const abs = cursor + i
      const marked = abs >= start && abs < end
      let j = i + 1
      while (j < chunk.length) {
        const a = cursor + j
        const m = a >= start && a < end
        if (m !== marked) break
        j++
      }
      const text = chunk.slice(i, j)
      const nextProps = marked ? { ...run.props, ...mark } : { ...run.props }
      // Normalize toggled-off boolean marks to absent
      for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
        if (nextProps[key] === false) delete nextProps[key]
      }
      pieces.push({
        kind: 'text',
        text,
        props: nextProps,
        ...(run.styleId !== undefined ? { styleId: run.styleId } : {}),
      })
      i = j
    }
    cursor += chunk.length
  }

  const runs: Run[] = pieces.map((p) => {
    if (p.kind === 'atomic') return { ...p.run, id: nextId('r') }
    return {
      id: nextId('r'),
      props: p.props,
      ...(p.styleId !== undefined ? { styleId: p.styleId } : {}),
      content: { type: 'text' as const, text: p.text },
    }
  })

  return { ...paragraph, runs: coalesceRuns(runs) }
}
