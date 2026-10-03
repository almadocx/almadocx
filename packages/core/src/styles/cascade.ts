import type {
  CharacterProps,
  Document,
  Paragraph,
  ParagraphProps,
  Run,
  StyleId,
} from '../model/types.js'

function mergeChar(base: CharacterProps, overlay: CharacterProps): CharacterProps {
  return {
    ...base,
    ...stripUndefined(overlay),
  }
}

function mergePara(base: ParagraphProps, overlay: ParagraphProps): ParagraphProps {
  return {
    ...base,
    ...stripUndefined(overlay),
  }
}

function stripUndefined<T extends object>(obj: T): Partial<T> {
  const out: Partial<T> = {}
  for (const [k, v] of Object.entries(obj) as [keyof T, T[keyof T]][]) {
    if (v !== undefined) out[k] = v
  }
  return out
}

function resolveParagraphStyleChain(doc: Document, styleId: StyleId | undefined): ParagraphProps {
  const seen = new Set<string>()
  let current = styleId
  const chain: ParagraphProps[] = []
  while (current && !seen.has(current)) {
    seen.add(current)
    const style = doc.styles.paragraphStyles[current]
    if (!style) break
    chain.push(style.paragraph)
    current = style.basedOn
  }
  let merged: ParagraphProps = {}
  for (let i = chain.length - 1; i >= 0; i--) {
    const layer = chain[i]
    if (layer) merged = mergePara(merged, layer)
  }
  return merged
}

function resolveParagraphCharStyleChain(
  doc: Document,
  styleId: StyleId | undefined,
): CharacterProps {
  const seen = new Set<string>()
  let current = styleId
  const chain: CharacterProps[] = []
  while (current && !seen.has(current)) {
    seen.add(current)
    const style = doc.styles.paragraphStyles[current]
    if (!style) break
    chain.push(style.character)
    current = style.basedOn
  }
  let merged: CharacterProps = {}
  for (let i = chain.length - 1; i >= 0; i--) {
    const layer = chain[i]
    if (layer) merged = mergeChar(merged, layer)
  }
  return merged
}

function resolveCharacterStyleChain(doc: Document, styleId: StyleId | undefined): CharacterProps {
  const seen = new Set<string>()
  let current = styleId
  const chain: CharacterProps[] = []
  while (current && !seen.has(current)) {
    seen.add(current)
    const style = doc.styles.characterStyles[current]
    if (!style) break
    chain.push(style.props)
    current = style.basedOn
  }
  let merged: CharacterProps = {}
  for (let i = chain.length - 1; i >= 0; i--) {
    const layer = chain[i]
    if (layer) merged = mergeChar(merged, layer)
  }
  return merged
}

/**
 * Document defaults → paragraph style → paragraph direct → character style → run direct.
 * Styles are never flattened on import; this resolves at read/layout time.
 */
function effectiveParagraphStyleId(doc: Document, paragraph: Paragraph): StyleId | undefined {
  return paragraph.props.styleId ?? doc.styles.defaultParagraphStyle ?? 'Normal'
}

export function resolveParagraphProps(doc: Document, paragraph: Paragraph): ParagraphProps {
  const styleId = effectiveParagraphStyleId(doc, paragraph)
  const fromStyle = resolveParagraphStyleChain(doc, styleId)
  return mergePara(mergePara(doc.styles.docDefaults.paragraph, fromStyle), paragraph.props)
}

export function resolveRunProps(doc: Document, paragraph: Paragraph, run: Run): CharacterProps {
  const styleId = effectiveParagraphStyleId(doc, paragraph)
  const paraStyleChar = resolveParagraphCharStyleChain(doc, styleId)
  const charStyle = resolveCharacterStyleChain(doc, run.styleId)
  return mergeChar(
    mergeChar(mergeChar(doc.styles.docDefaults.character, paraStyleChar), charStyle),
    run.props,
  )
}
