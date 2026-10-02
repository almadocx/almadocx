import { strFromU8, unzipSync, zipSync, strToU8 } from 'fflate'
import { AlmadocxError, assert } from '../util/assert.js'
import { PACKAGE_LIMITS } from './limits.js'

export type ZipEntries = Map<string, Uint8Array>

function normalizePath(name: string): string {
  const n = name.replace(/\\/g, '/')
  assert(n.length > 0, 'zip_path', 'empty entry name')
  assert(n.length <= PACKAGE_LIMITS.maxPathLength, 'zip_path', 'entry name too long')
  assert(!n.startsWith('/'), 'zip_path', `absolute path rejected: ${n}`)
  assert(!/^[a-zA-Z]:/.test(n), 'zip_path', `drive path rejected: ${n}`)
  const parts = n.split('/')
  for (const part of parts) {
    assert(part !== '..', 'zip_path', `path traversal rejected: ${n}`)
    assert(part !== '.', 'zip_path', `invalid path segment: ${n}`)
  }
  return n
}

export function readZip(bytes: Uint8Array): ZipEntries {
  assert(
    bytes.byteLength <= PACKAGE_LIMITS.maxCompressedBytes,
    'zip_size',
    'package exceeds compressed size limit',
  )

  let raw: Record<string, Uint8Array>
  try {
    raw = unzipSync(bytes, {
      filter: (file) => {
        if (file.originalSize > PACKAGE_LIMITS.maxUncompressedEntryBytes) {
          throw new AlmadocxError('zip_bomb', `entry too large: ${file.name}`)
        }
        return true
      },
    })
  } catch (e) {
    if (e instanceof AlmadocxError) throw e
    throw new AlmadocxError('zip_corrupt', `failed to read package: ${String(e)}`)
  }

  const names = Object.keys(raw)
  assert(names.length <= PACKAGE_LIMITS.maxEntries, 'zip_entries', 'too many zip entries')

  const out: ZipEntries = new Map()
  let total = 0
  for (const name of names) {
    // unzipSync always yields a Uint8Array for each key from Object.keys
    const data = raw[name]!
    if (name.endsWith('/')) continue
    const path = normalizePath(name)
    total += data.byteLength
    assert(
      total <= PACKAGE_LIMITS.maxTotalUncompressedBytes,
      'zip_bomb',
      'total uncompressed size exceeds limit',
    )
    out.set(path, data)
  }

  // Compression ratio check against original package size
  if (bytes.byteLength > 0) {
    const ratio = total / bytes.byteLength
    assert(
      ratio <= PACKAGE_LIMITS.maxCompressionRatio,
      'zip_bomb',
      `suspicious compression ratio: ${ratio.toFixed(1)}`,
    )
  }

  return out
}

export function writeZip(entries: ZipEntries): Uint8Array {
  const record: Record<string, Uint8Array> = {}
  for (const [path, data] of entries) {
    record[normalizePath(path)] = data
  }
  return zipSync(record, { level: 6 })
}

export function zipText(entries: ZipEntries, path: string): string | undefined {
  const data = entries.get(path)
  if (!data) return undefined
  assert(data.byteLength <= PACKAGE_LIMITS.maxXmlBytes, 'xml_size', `${path} too large`)
  return strFromU8(data)
}

export function setZipText(entries: ZipEntries, path: string, text: string): void {
  entries.set(normalizePath(path), strToU8(text))
}

export function isMacroPath(path: string): boolean {
  const p = path.toLowerCase()
  return (
    p.includes('vba') ||
    p.includes('macrosheets') ||
    (p.endsWith('.bin') && p.includes('macro')) ||
    p.includes('basic/') ||
    (p.endsWith('module.xml') && p.includes('basic'))
  )
}
