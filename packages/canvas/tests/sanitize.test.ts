import { describe, expect, it } from 'vitest'
import { htmlToFragment, sanitizeHtml } from '../src/clipboard/sanitize.js'

describe('clipboard sanitize', () => {
  it('strips script tags', () => {
    const html = sanitizeHtml('<p>Hi<script>alert(1)</script></p>')
    expect(html.toLowerCase()).not.toContain('script')
    expect(html).toContain('Hi')
  })

  it('maps bold/italic from html', () => {
    const frag = htmlToFragment('<p>Hello <strong>bold</strong> <em>italic</em></p>')
    expect(frag.paragraphs.length).toBeGreaterThanOrEqual(1)
    const runs = frag.paragraphs.flatMap((p) => p.runs)
    expect(runs.some((r) => r.content.type === 'text' && r.props.bold)).toBe(true)
    expect(runs.some((r) => r.content.type === 'text' && r.props.italic)).toBe(true)
  })

  it('strips event handlers', () => {
    const html = sanitizeHtml('<p onclick="evil()">x</p>')
    expect(html.toLowerCase()).not.toContain('onclick')
  })
})
