import { describe, expect, it } from 'vitest'
import {
  createEmptyDocument,
  createEmptyParagraph,
  createTextRun,
  type Document,
  type Table,
} from '@almadocx/core'
import { A11yMirror } from '../src/a11y/mirror.js'

function para(text: string, props: Record<string, unknown> = {}) {
  const p = createEmptyParagraph()
  p.props = props as typeof p.props
  p.runs = text ? [createTextRun(text)] : []
  return p
}

function makeTable(): Table {
  return {
    id: 't1',
    type: 'table',
    props: {},
    rows: [
      {
        id: 'r0',
        props: {},
        cells: [
          {
            id: 'c00',
            props: { gridSpan: 2 },
            blocks: [para('A')],
          },
          {
            id: 'c01',
            props: {},
            blocks: [para('B')],
          },
        ],
      },
      {
        id: 'r1',
        props: {},
        cells: [
          {
            id: 'c10',
            props: { vMerge: 'continue' },
            blocks: [para('hidden')],
          },
          {
            id: 'c11',
            props: { vMerge: 'restart' },
            blocks: [para('C'), para('D')],
          },
        ],
      },
    ],
  }
}

function richDoc(): Document {
  const doc = createEmptyDocument()
  const section = doc.sections[0]!
  section.header = {
    id: 'h1',
    blocks: [para('Header text'), para('')],
  }
  section.footer = {
    id: 'f1',
    blocks: [para(''), para('Footer text')],
  }
  const imagePara = createEmptyParagraph()
  imagePara.runs = [
    {
      id: 'img1',
      props: {},
      content: {
        type: 'image',
        mediaId: 'm1',
        widthTwips: 100,
        heightTwips: 100,
        alt: 'Logo',
      },
    },
    createTextRun(' after'),
    {
      id: 'img2',
      props: {},
      content: {
        type: 'image',
        mediaId: 'm2',
        widthTwips: 50,
        heightTwips: 50,
      },
    },
  ]
  section.blocks = [
    para('Intro'),
    para('Bullet', { numPr: { numId: '1', ilvl: 0 } }),
    para('Bullet2', { numPr: { numId: '1', ilvl: 0 } }),
    para('Numbered', { numPr: { numId: '2', ilvl: 0 } }),
    para(''),
    makeTable(),
    imagePara,
    para('Trail', { numPr: { numId: '1', ilvl: 0 } }),
  ]
  return doc
}

describe('A11yMirror', () => {
  it('builds an ARIA document tree with sections, lists, and tables', () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const mirror = new A11yMirror(host)

    expect(mirror.root.getAttribute('role')).toBe('document')
    expect(host.contains(mirror.root)).toBe(true)

    mirror.sync(richDoc(), 'caret at 3')

    expect(mirror.root.getAttribute('aria-description')).toBe('caret at 3')
    const region = mirror.root.querySelector('[role="region"]')
    expect(region).toBeTruthy()
    expect(region!.querySelector('header')?.getAttribute('aria-label')).toBe('Header')
    expect(region!.querySelector('footer')?.getAttribute('aria-label')).toBe('Footer')
    expect(region!.querySelectorAll('ul li').length).toBeGreaterThanOrEqual(2)
    expect(region!.querySelectorAll('ol li').length).toBe(1)

    const grid = region!.querySelector('[role="grid"]')!
    expect(grid.getAttribute('aria-label')).toBe('Table')
    expect(grid.getAttribute('aria-rowcount')).toBe('2')
    expect(grid.getAttribute('aria-colcount')).toBe('2')

    const cells = Array.from(grid.querySelectorAll('[role="gridcell"]'))
    expect(cells.some((c) => c.getAttribute('aria-colspan') === '2')).toBe(true)
    expect(cells.map((c) => c.textContent).join('|')).not.toContain('hidden')
    expect(cells.some((c) => c.textContent === 'C D')).toBe(true)

    const paragraphs = region!.querySelectorAll('[role="paragraph"]')
    expect(
      Array.from(paragraphs).some((p) => p.textContent === 'Logo afterImage'),
    ).toBe(true)
    expect(Array.from(paragraphs).some((p) => p.textContent === '\u00a0')).toBe(true)

    mirror.destroy()
    expect(host.contains(mirror.root)).toBe(false)
    host.remove()
  })

  it('syncs without caret announcement and handles empty tables', () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const mirror = new A11yMirror(host)
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [
      {
        id: 't-empty',
        type: 'table',
        props: {},
        rows: [],
      },
      para('Only'),
    ]
    mirror.sync(doc)
    expect(mirror.root.hasAttribute('aria-description')).toBe(false)
    const grid = mirror.root.querySelector('[role="grid"]')!
    expect(grid.getAttribute('aria-rowcount')).toBe('0')
    expect(grid.getAttribute('aria-colcount')).toBe('0')
    mirror.destroy()
    host.remove()
  })

  it('uses nbsp for empty table cells and empty list items', () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const mirror = new A11yMirror(host)
    const doc = createEmptyDocument()
    doc.sections[0]!.blocks = [
      para('', { numPr: { numId: '1', ilvl: 0 } }),
      {
        id: 't-blank',
        type: 'table',
        props: {},
        rows: [
          {
            id: 'r0',
            props: {},
            cells: [{ id: 'c0', props: {}, blocks: [para('')] }],
          },
        ],
      },
    ]
    mirror.sync(doc)
    const li = mirror.root.querySelector('ul li')!
    expect(li.textContent).toBe('\u00a0')
    const cell = mirror.root.querySelector('[role="gridcell"]')!
    expect(cell.textContent).toBe('\u00a0')
    mirror.destroy()
    host.remove()
  })
})
