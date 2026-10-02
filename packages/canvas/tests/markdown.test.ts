import { describe, expect, it } from 'vitest'
import { markdownToFragment } from '../src/clipboard/markdown.js'

describe('markdownToFragment', () => {
  it('returns undefined when text has no markdown signals', () => {
    expect(markdownToFragment('plain text only')).toBeUndefined()
  })

  it('parses headings, lists, inline marks, and plain lines', () => {
    const frag = markdownToFragment(
      [
        '# Title',
        '## Sub',
        '- bullet **bold** and *italic*',
        '* also bullet',
        '+ plus',
        '1. numbered __strong__',
        '2) alt numbered `code`',
        'trailing _em_ text',
        '',
      ].join('\n'),
    )
    expect(frag).toBeDefined()
    expect(frag!.paragraphs.length).toBeGreaterThan(5)
    expect(frag!.paragraphs[0]!.props.styleId).toBe('Heading1')
    expect(frag!.paragraphs[1]!.props.styleId).toBe('Heading2')

    const bullets = frag!.paragraphs.filter((p) => p.props.numPr?.numId === '1')
    expect(bullets.length).toBe(3)
    const numbered = frag!.paragraphs.filter((p) => p.props.numPr?.numId === '2')
    expect(numbered.length).toBe(2)

    const boldRuns = frag!.paragraphs.flatMap((p) => p.runs).filter((r) => r.props.bold)
    const italicRuns = frag!.paragraphs.flatMap((p) => p.runs).filter((r) => r.props.italic)
    const codeRuns = frag!.paragraphs
      .flatMap((p) => p.runs)
      .filter((r) => r.props.fontFamily === 'Consolas')
    expect(boldRuns.length).toBeGreaterThan(0)
    expect(italicRuns.length).toBeGreaterThan(0)
    expect(codeRuns.length).toBe(1)
  })

  it('normalizes CRLF and lone CR, and handles empty inline parse', () => {
    const frag = markdownToFragment('**only**\r\nline\r# H')
    expect(frag).toBeDefined()
    expect(frag!.paragraphs.some((p) => p.props.styleId === 'Heading1')).toBe(true)

    const emptyish = markdownToFragment('# \n- \n1. ')
    expect(emptyish).toBeDefined()
    expect(emptyish!.paragraphs.length).toBe(3)
    expect(emptyish!.paragraphs.every((p) => p.runs.length >= 1)).toBe(true)
  })
})
