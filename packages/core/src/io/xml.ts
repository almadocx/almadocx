import { XMLParser, XMLBuilder } from 'fast-xml-parser'
import { AlmadocxError, assert } from '../util/assert.js'
import { PACKAGE_LIMITS } from './limits.js'

/**
 * XXE-safe parse: no DTD / external entity processing.
 * fast-xml-parser does not resolve external entities by default when processEntities is controlled.
 */
export function parseXml(xml: string): unknown {
  assert(xml.length <= PACKAGE_LIMITS.maxXmlBytes, 'xml_size', 'XML exceeds size limit')
  // Reject DOCTYPE to mitigate XXE / billion laughs via DTD
  if (/<!DOCTYPE/i.test(xml) || /<!ENTITY/i.test(xml)) {
    throw new AlmadocxError('xml_xxe', 'DTD/ENTITY declarations are not allowed')
  }

  try {
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      textNodeName: '#text',
      preserveOrder: false,
      trimValues: false,
      processEntities: false,
      allowBooleanAttributes: true,
      removeNSPrefix: false,
    })
    return parser.parse(xml)
  } catch (e) {
    throw new AlmadocxError('xml_parse', `XML parse failed: ${String(e)}`)
  }
}

export function buildXml(obj: unknown): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    format: true,
    suppressEmptyNode: false,
    processEntities: false,
  })
  return String(builder.build(obj as never))
}

export function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return []
  return Array.isArray(value) ? value : [value]
}

export function xmlAttr(node: Record<string, unknown> | undefined, name: string): string | undefined {
  if (!node) return undefined
  const v = node[`@_${name}`]
  return typeof v === 'string' || typeof v === 'number' ? String(v) : undefined
}

export function xmlText(node: unknown): string {
  if (node === undefined || node === null) return ''
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  if (typeof node === 'object') {
    const rec = node as Record<string, unknown>
    if (typeof rec['#text'] === 'string' || typeof rec['#text'] === 'number') {
      return String(rec['#text'])
    }
  }
  return ''
}
