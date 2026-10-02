import { describe, expect, it } from 'vitest'
import { zipSync, strToU8 } from 'fflate'
import { AlmadocxError, parseXml, readZip } from '../src/index.js'

describe('security', () => {
  it('rejects path traversal in zip entries', () => {
    const bytes = zipSync({ '../evil.txt': strToU8('x') })
    expect(() => readZip(bytes)).toThrow(AlmadocxError)
  })

  it('rejects DTD/ENTITY in XML', () => {
    const xml = `<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><root>&xxe;</root>`
    expect(() => parseXml(xml)).toThrow(AlmadocxError)
  })

  it('rejects absolute zip paths', () => {
    const bytes = zipSync({ '/tmp/x.txt': strToU8('x') })
    expect(() => readZip(bytes)).toThrow(AlmadocxError)
  })
})
