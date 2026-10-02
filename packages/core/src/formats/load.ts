import type { Document, SourceFormat } from '../model/types.js'
import { AlmadocxError } from '../util/assert.js'
import { parseDocx, sniffDocx } from './docx/parse.js'
import { serializeDocx } from './docx/serialize.js'
import { parseOdt, sniffOdt } from './odt/parse.js'
import { serializeOdt } from './odt/serialize.js'

export function detectFormat(bytes: Uint8Array): SourceFormat | undefined {
  // ODT has mimetype; check both carefully without double-throw
  const odt = sniffOdt(bytes)
  const docx = sniffDocx(bytes)
  if (odt && !docx) return 'odt'
  if (docx && !odt) return 'docx'
  if (odt && docx) {
    // Prefer explicit mimetype path — sniffOdt requires mimetype
    return 'odt'
  }
  return undefined
}

export function loadDocument(bytes: Uint8Array, formatHint?: SourceFormat): Document {
  const format = formatHint ?? detectFormat(bytes)
  if (format === 'docx') return parseDocx(bytes)
  if (format === 'odt') return parseOdt(bytes)
  throw new AlmadocxError('format_unknown', 'Could not detect DOCX or ODT package')
}

export function saveDocument(doc: Document, format?: SourceFormat): Uint8Array {
  const target = format ?? (doc.package.sourceFormat === 'internal' ? 'docx' : doc.package.sourceFormat)
  if (target === 'docx' || target === 'internal') {
    const out = serializeDocx({
      ...doc,
      package: { ...doc.package, sourceFormat: 'docx' },
    })
    return out
  }
  if (target === 'odt') {
    return serializeOdt({
      ...doc,
      package: { ...doc.package, sourceFormat: 'odt' },
    })
  }
  throw new AlmadocxError('format_unknown', `Cannot save as ${String(target)}`)
}
