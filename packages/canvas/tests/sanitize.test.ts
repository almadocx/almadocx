import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  createEmptyParagraph,
  createTextRun,
  type Paragraph,
  type Run,
} from '@almadocx/core'
import {
  htmlToFragment,
  plainTextToFragment,
  runsToHtml,
  sanitizeHtml,
} from '../src/clipboard/sanitize.js'

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

  it('sanitizes style / font / lists / headings and unwraps unknown tags', () => {
    const html = sanitizeHtml(`
      <style>.x{}</style>
      <iframe src="x"></iframe>
      <object></object>
      <embed></embed>
      <div>
        <span style="font-weight:bold;font-style:italic;text-decoration:underline line-through;font-size:16px;font-family:'Arial',sans-serif;color:#abc;background-color:rgb(1,2,3)">styled</span>
        <span style="text-decoration:strike;color:;bogus;font-size:NaNpt;font-family: ;background:not-a-color">more</span>
        <b>b</b><i>i</i><u>u</u><s>s</s><strike>k</strike><del>d</del>
        <sub>sub</sub><sup>sup</sup>
        <font face="Courier" color="#ff0000" size="3">font</font>
        <a href="http://evil">link</a>
        <!-- comment -->
      </div>
      <h1>H1</h1>
      <ul><li>item</li></ul>
      <br/>
      <p style="font-weight:700;font-size:12pt;color:not-a-color;background:#fff">p</p>
    `)
    expect(html.toLowerCase()).not.toContain('iframe')
    expect(html.toLowerCase()).not.toContain('href')
    expect(html).toContain('styled')

    const frag = htmlToFragment(html)
    expect(frag.paragraphs.some((p) => p.props.styleId === 'Heading1')).toBe(true)
    expect(frag.paragraphs.some((p) => p.props.numPr?.numId === '1')).toBe(true)
    const runs = frag.paragraphs.flatMap((p) => p.runs)
    expect(runs.some((r) => r.content.type === 'break')).toBe(true)
    expect(runs.some((r) => r.props.verticalAlign === 'subscript')).toBe(true)
    expect(runs.some((r) => r.props.verticalAlign === 'superscript')).toBe(true)
    expect(runs.some((r) => r.props.fontFamily === 'Courier')).toBe(true)
    expect(runs.some((r) => r.props.highlight)).toBe(true)
    expect(runs.some((r) => r.props.strike)).toBe(true)
  })

  it('handles empty html and plain text fragments', () => {
    const empty = htmlToFragment('')
    expect(empty.paragraphs.length).toBe(1)
    const plain = plainTextToFragment('a\r\nb\rc')
    expect(plain.paragraphs.map((p) => (p.runs[0]!.content as { text: string }).text)).toEqual([
      'a',
      'b',
      'c',
    ])
  })

  it('covers empty block tags, missing root, and non-element walk nodes', () => {
    const emptyBlocks = htmlToFragment('<p></p><div></div><h2></h2>')
    expect(emptyBlocks.paragraphs.length).toBeGreaterThanOrEqual(2)
    expect(emptyBlocks.paragraphs.every((p) => p.runs.length >= 1)).toBe(true)

    const rootSpy = vi.spyOn(Document.prototype, 'getElementById').mockReturnValue(null)
    expect(sanitizeHtml('<p>x</p>')).toBe('')
    rootSpy.mockRestore()

    const original = DOMParser.prototype.parseFromString
    let calls = 0
    DOMParser.prototype.parseFromString = function (
      this: DOMParser,
      str: string,
      type: DOMParserSupportedType,
    ) {
      const doc = original.call(this, str, type)
      calls += 1
      if (calls >= 2) {
        const root = doc.getElementById('root')
        root?.appendChild(doc.createComment('skip-me'))
        root?.appendChild(doc.createTextNode(''))
      }
      return doc
    }
    try {
      const frag = htmlToFragment('<p>ok</p>')
      expect(
        frag.paragraphs.some((p) =>
          p.runs.some((r) => r.content.type === 'text' && r.content.text === 'ok'),
        ),
      ).toBe(true)
    } finally {
      DOMParser.prototype.parseFromString = original
    }
  })

  it('serializes runs to html including marks and non-text content', () => {
    const p: Paragraph = createEmptyParagraph()
    p.runs = [
      createTextRun('&<>', { bold: true, italic: true, underline: true, strike: true }),
      { id: 't', props: {}, content: { type: 'tab' } },
      { id: 'br', props: {}, content: { type: 'break', breakType: 'line' } },
      {
        id: 'img',
        props: {},
        content: { type: 'image', mediaId: 'm', widthTwips: 1, heightTwips: 1 },
      },
    ]
    const emptyP = createEmptyParagraph()
    emptyP.runs = []
    const html = runsToHtml([p, emptyP])
    expect(html).toContain('<strong>')
    expect(html).toContain('&amp;')
    expect(html).toContain('&emsp;')
    expect(html).toContain('<br/>')
    expect(html).toContain('<p><br/></p>')
  })
})

describe('sanitizeHtml without document', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('uses the node fallback path when document is undefined', () => {
    const realDocument = globalThis.document
    vi.stubGlobal('document', undefined)
    const html = sanitizeHtml('<p onclick="x">Hi<script>bad</script><style>z</style></p>')
    expect(html.toLowerCase()).not.toContain('script')
    expect(html.toLowerCase()).not.toContain('onclick')
    expect(html).toContain('Hi')

    // htmlToFragment also takes the no-document branch
    const frag = htmlToFragment('<p>a</p>\n<b>b</b>')
    expect(frag.paragraphs.length).toBeGreaterThanOrEqual(1)
    vi.stubGlobal('document', realDocument)
  })
})
