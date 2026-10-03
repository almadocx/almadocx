import { XMLParser, XMLBuilder } from 'fast-xml-parser'
import { AlmadocxError, assert } from '../util/assert.js'
import { PACKAGE_LIMITS } from './limits.js'

/**
 * XXE-safe parse: no DTD / external entity processing.
 * fast-xml-parser does not resolve external entities by default when processEntities is controlled.
 */
const XML_PARSER_BASE = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  trimValues: false,
  processEntities: false,
  allowBooleanAttributes: true,
  removeNSPrefix: false,
} as const

function assertSafeXml(xml: string): void {
  assert(xml.length <= PACKAGE_LIMITS.maxXmlBytes, 'xml_size', 'XML exceeds size limit')
  // Reject DOCTYPE to mitigate XXE / billion laughs via DTD
  if (/<!DOCTYPE/i.test(xml) || /<!ENTITY/i.test(xml)) {
    throw new AlmadocxError('xml_xxe', 'DTD/ENTITY declarations are not allowed')
  }
}

export function parseXml(xml: string): unknown {
  assertSafeXml(xml)
  try {
    return new XMLParser({ ...XML_PARSER_BASE, preserveOrder: false }).parse(xml)
  } catch (e) {
    throw new AlmadocxError('xml_parse', `XML parse failed: ${String(e)}`)
  }
}

/** Ordered parse so sibling `w:p` / `w:tbl` keep document order (body blocks). */
export function parseXmlOrdered(xml: string): unknown {
  assertSafeXml(xml)
  try {
    return new XMLParser({ ...XML_PARSER_BASE, preserveOrder: true }).parse(xml)
  } catch (e) {
    throw new AlmadocxError('xml_parse', `XML parse failed: ${String(e)}`)
  }
}

/**
 * Collapse a preserveOrder element-children array into a classic object tree
 * (repeated tags become arrays). Suitable for feeding existing findChild parsers.
 *
 * Attributes on a child (`:@`) are attached to that child object — never flattened
 * onto the parent (which would collide e.g. multiple `@_w:val` on `tcPr`).
 */
export function collapseOrdered(nodes: unknown): Record<string, unknown> {
  if (!Array.isArray(nodes)) {
    return nodes && typeof nodes === 'object' ? (nodes as Record<string, unknown>) : {}
  }
  const out: Record<string, unknown> = {}
  for (const raw of nodes) {
    if (!raw || typeof raw !== 'object') continue
    const item = raw as Record<string, unknown>
    const attrs = item[':@']
    for (const [k, v] of Object.entries(item)) {
      if (k === ':@') continue
      let child: Record<string, unknown> | string | number | boolean
      if (Array.isArray(v)) {
        child = collapseOrdered(v)
      } else if (v && typeof v === 'object') {
        child = { ...(v as Record<string, unknown>) }
      } else if (v === undefined) {
        child = {}
      } else {
        child = v as string | number | boolean
      }
      if (attrs && typeof attrs === 'object' && typeof child === 'object') {
        child = { ...attrs, ...child }
      }
      if (out[k] === undefined) out[k] = child
      else out[k] = [...asArray(out[k] as never), child as never]
    }
  }
  return out
}

/** One preserveOrder sibling: `{ "w:p": [...], ":@": {...} }` → local name + collapsed node. */
export function orderedElement(
  item: unknown,
): { name: string; node: Record<string, unknown> } | undefined {
  if (!item || typeof item !== 'object') return undefined
  const rec = item as Record<string, unknown>
  const tag = Object.keys(rec).find((k) => k !== ':@')
  if (!tag) return undefined
  const children = rec[tag]
  const node = Array.isArray(children) ? collapseOrdered(children) : {}
  const attrs = rec[':@']
  if (attrs && typeof attrs === 'object') Object.assign(node, attrs)
  const colon = tag.indexOf(':')
  const name = colon >= 0 ? tag.slice(colon + 1) : tag
  return { name, node }
}

/** Decode common XML character entities without enabling external entity resolution. */
export function decodeXmlEntities(text: string): string {
  return text.replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, (m, body: string) => {
    if (body === 'amp') return '&'
    if (body === 'lt') return '<'
    if (body === 'gt') return '>'
    if (body === 'quot') return '"'
    if (body === 'apos') return "'"
    if (body.startsWith('#x') || body.startsWith('#X')) {
      const cp = Number.parseInt(body.slice(2), 16)
      return Number.isFinite(cp) ? String.fromCodePoint(cp) : m
    }
    if (body.startsWith('#')) {
      const cp = Number(body.slice(1))
      return Number.isFinite(cp) ? String.fromCodePoint(cp) : m
    }
    return m
  })
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

/** Try attribute names in order (e.g. `w:val` then unprefixed `val`). */
export function xmlAttrAlt(
  node: Record<string, unknown> | undefined,
  ...names: string[]
): string | undefined {
  for (const name of names) {
    const v = xmlAttr(node, name)
    if (v !== undefined) return v
  }
  return undefined
}

/** Like {@link xmlAttrAlt} but returns `fallback` when none match. */
export function xmlAttrAltOr(
  node: Record<string, unknown> | undefined,
  fallback: string,
  ...names: string[]
): string {
  return xmlAttrAlt(node, ...names) ?? fallback
}

export function xmlText(node: unknown): string {
  if (node === undefined || node === null) return ''
  if (typeof node === 'string' || typeof node === 'number') return decodeXmlEntities(String(node))
  if (typeof node === 'object') {
    const rec = node as Record<string, unknown>
    if (typeof rec['#text'] === 'string' || typeof rec['#text'] === 'number') {
      return decodeXmlEntities(String(rec['#text']))
    }
  }
  return ''
}
