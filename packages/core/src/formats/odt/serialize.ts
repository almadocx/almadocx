import type { CharacterProps, Document, Paragraph, ParagraphProps, Run } from '../../model/types.js'
import { setZipText, writeZip, type ZipEntries } from '../../io/zip.js'
import { twipsToHm } from '../../util/units.js'
import { AlmadocxError } from '../../util/assert.js'

function esc(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function hm(twips: number | undefined): string | undefined {
  if (twips === undefined) return undefined
  return `${(twipsToHm(twips) / 100).toFixed(3)}mm`
}

function textStyleAttrs(props: CharacterProps): string {
  const attrs: string[] = []
  if (props.bold) attrs.push('fo:font-weight="bold"')
  if (props.italic) attrs.push('fo:font-style="italic"')
  if (props.underline) attrs.push('style:text-underline-style="solid"')
  if (props.strike) attrs.push('style:text-line-through-style="solid"')
  if (props.fontSizePt !== undefined) attrs.push(`fo:font-size="${String(props.fontSizePt)}pt"`)
  if (props.color) attrs.push(`fo:color="${esc(props.color)}"`)
  if (props.fontFamily) attrs.push(`style:font-name="${esc(props.fontFamily)}"`)
  return attrs.join(' ')
}

function paraStyleAttrs(props: ParagraphProps): string {
  const attrs: string[] = []
  if (props.alignment) {
    const map = { left: 'start', right: 'end', center: 'center', justify: 'justify' } as const
    attrs.push(`fo:text-align="${map[props.alignment]}"`)
  }
  const ml = hm(props.indentLeft)
  const mr = hm(props.indentRight)
  const mt = hm(props.spacingBefore)
  const mb = hm(props.spacingAfter)
  const ti = hm(props.indentFirstLine)
  if (ml) attrs.push(`fo:margin-left="${ml}"`)
  if (mr) attrs.push(`fo:margin-right="${mr}"`)
  if (mt) attrs.push(`fo:margin-top="${mt}"`)
  if (mb) attrs.push(`fo:margin-bottom="${mb}"`)
  if (ti) attrs.push(`fo:text-indent="${ti}"`)
  if (props.lineSpacing !== undefined && props.lineSpacingRule === 'auto') {
    attrs.push(`fo:line-height="${String(Math.round(props.lineSpacing * 100))}%"`)
  }
  return attrs.join(' ')
}

function runToOdt(run: Run, autoStyles: Map<string, string>, counter: { n: number }): string {
  switch (run.content.type) {
    case 'tab':
      return '<text:tab/>'
    case 'break':
      return '<text:line-break/>'
    case 'image':
      // Image binary embedding is handled at package level in a later pass; emit a text fallback.
      return esc(run.content.alt ?? '[image]')
    case 'text': {
      const hasProps = Object.keys(run.props).length > 0
      const text = esc(run.content.text)
      if (!hasProps) return text
      const key = JSON.stringify(run.props)
      let styleName = autoStyles.get(key)
      if (!styleName) {
        counter.n += 1
        styleName = `T${String(counter.n)}`
        autoStyles.set(key, styleName)
      }
      return `<text:span text:style-name="${styleName}">${text}</text:span>`
    }
  }
}

function paragraphToOdt(
  p: Paragraph,
  autoStyles: Map<string, string>,
  paraStyles: Map<string, string>,
  counter: { n: number },
): string {
  let styleName = p.props.styleId
  const paraKey = JSON.stringify(p.props)
  if (!styleName || Object.keys(p.props).some((k) => k !== 'styleId')) {
    let existing = paraStyles.get(paraKey)
    if (!existing) {
      counter.n += 1
      existing = `P${String(counter.n)}`
      paraStyles.set(paraKey, existing)
    }
    styleName = existing
  }
  const inner = p.runs.map((r) => runToOdt(r, autoStyles, counter)).join('')
  return `<text:p text:style-name="${esc(styleName!)}">${inner}</text:p>`
}

function buildAutomaticStyles(
  autoStyles: Map<string, string>,
  paraStyles: Map<string, string>,
): string {
  const parts: string[] = []
  for (const [key, name] of paraStyles) {
    const props = JSON.parse(key) as ParagraphProps
    const attrs = paraStyleAttrs(props)
    parts.push(
      `<style:style style:name="${esc(name)}" style:family="paragraph">` +
        `<style:paragraph-properties ${attrs}/>` +
        `</style:style>`,
    )
  }
  for (const [key, name] of autoStyles) {
    const props = JSON.parse(key) as CharacterProps
    const attrs = textStyleAttrs(props)
    parts.push(
      `<style:style style:name="${esc(name)}" style:family="text">` +
        `<style:text-properties ${attrs}/>` +
        `</style:style>`,
    )
  }
  return parts.join('')
}

export function serializeOdt(doc: Document): Uint8Array {
  const autoStyles = new Map<string, string>()
  const paraStyles = new Map<string, string>()
  const counter = { n: 0 }

  const body = doc.sections
    .flatMap((s) => s.blocks)
    .filter((b): b is Paragraph => b.type === 'paragraph')
    .map((p) => paragraphToOdt(p, autoStyles, paraStyles, counter))
    .join('')

  const automatic = buildAutomaticStyles(autoStyles, paraStyles)

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
    `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
    `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
    `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
    `office:version="1.3">` +
    `<office:automatic-styles>${automatic}</office:automatic-styles>` +
    `<office:body><office:text>${body}</office:text></office:body>` +
    `</office:document-content>`

  const entries: ZipEntries = new Map()
  for (const part of doc.package.preservedParts) {
    if (part.path === 'content.xml') continue
    entries.set(part.path, part.bytes)
  }

  setZipText(entries, 'mimetype', 'application/vnd.oasis.opendocument.text')
  // mimetype must be stored first uncompressed ideally; fflate zipSync may compress — acceptable for Phase 1
  setZipText(entries, 'content.xml', content)

  if (!entries.has('styles.xml')) {
    setZipText(
      entries,
      'styles.xml',
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.3">` +
        `<office:styles>` +
        `<style:style style:name="Standard" style:family="paragraph"/>` +
        `</office:styles></office:document-styles>`,
    )
  }
  if (!entries.has('META-INF/manifest.xml')) {
    setZipText(
      entries,
      'META-INF/manifest.xml',
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
        `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>` +
        `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
        `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
        `</manifest:manifest>`,
    )
  }

  try {
    return writeZip(entries)
  } catch (e) {
    throw new AlmadocxError('odt_serialize', `failed to write odt: ${String(e)}`)
  }
}
