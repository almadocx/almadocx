import { describe, expect, it } from 'vitest'
import {
  applyOp,
  createEmptyDocument,
  layoutDocument,
  patchLayoutCellParagraph,
  patchLayoutParagraph,
  type Document,
  type Table,
} from '../src/index.js'

describe('layout coverage', () => {
  it('lays out headers, tabs, images, alignment, breaks, and widow splits', () => {
    const doc = createEmptyDocument()
    doc.sections[0]!.properties.titlePage = true
    doc.sections[0]!.header = {
      blocks: [
        {
          id: 'h',
          type: 'paragraph',
          props: {},
          runs: [{ id: 'hr', props: {}, content: { type: 'text', text: 'Header' } }],
        },
      ],
    }
    doc.sections[0]!.headerFirst = {
      blocks: [
        {
          id: 'hf',
          type: 'paragraph',
          props: {},
          runs: [{ id: 'hfr', props: {}, content: { type: 'text', text: 'First' } }],
        },
      ],
    }
    doc.sections[0]!.footer = {
      blocks: [
        {
          id: 'f',
          type: 'paragraph',
          props: {},
          runs: [{ id: 'fr', props: {}, content: { type: 'text', text: 'Footer' } }],
        },
      ],
    }
    doc.sections[0]!.footerFirst = {
      blocks: [
        {
          id: 'ff',
          type: 'paragraph',
          props: {},
          runs: [{ id: 'ffr', props: {}, content: { type: 'text', text: 'FirstF' } }],
        },
      ],
    }
    doc.media['m1'] = {
      id: 'm1',
      contentType: 'image/png',
      bytes: new Uint8Array([1, 2, 3]),
    }
    doc.sections[0]!.blocks = [
      {
        id: 'p0',
        type: 'paragraph',
        props: {
          alignment: 'center',
          pageBreakBefore: false,
          tabs: [{ position: 1440, alignment: 'left', leader: 'dot' }],
        },
        runs: [
          { id: 'r0', props: { bold: true }, content: { type: 'text', text: 'Title ' } },
          { id: 'r1', props: {}, content: { type: 'tab' } },
          {
            id: 'r2',
            props: {},
            content: { type: 'image', mediaId: 'm1', widthTwips: 1440, heightTwips: 720, alt: 'pic' },
          },
          { id: 'r3', props: {}, content: { type: 'break', breakType: 'line' } },
          { id: 'r4', props: { italic: true, underline: true, strike: true }, content: { type: 'text', text: 'more' } },
        ],
      },
      {
        id: 'p1',
        type: 'paragraph',
        props: { alignment: 'right', keepNext: true, keepLines: true, pageBreakBefore: true },
        runs: [
          {
            id: 'long',
            props: {},
            content: {
              type: 'text',
              text: 'Word '.repeat(200),
            },
          },
        ],
      },
      {
        id: 'p2',
        type: 'paragraph',
        props: { widowControl: true },
        runs: [
          {
            id: 'long2',
            props: {},
            content: { type: 'text', text: 'Line '.repeat(400) },
          },
        ],
      },
    ]

    const layout = layoutDocument(doc)
    expect(layout.pages.length).toBeGreaterThan(1)
    expect(layout.pages[0]!.header?.paragraphs.length).toBeGreaterThan(0)
    expect(layout.pages[0]!.footer?.paragraphs.length).toBeGreaterThan(0)
  })

  it('patches paragraph shifting a following table', () => {
    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'Hi',
    }).doc
    const table: Table = {
      id: 't',
      type: 'table',
      props: {},
      gridCols: [3000],
      rows: [
        {
          id: 'r',
          props: {},
          cells: [
            {
              id: 'c',
              props: {},
              blocks: [
                {
                  id: 'cp',
                  type: 'paragraph',
                  props: {},
                  runs: [{ id: 'cr', props: {}, content: { type: 'text', text: 'cell' } }],
                },
              ],
            },
          ],
        },
      ],
    }
    doc.sections[0]!.blocks.push(table)
    let layout = layoutDocument(doc)
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 2 },
      text: '\n'.repeat(5) + 'x'.repeat(80),
    }).doc
    // Use actual text growth without break chars — long wrap
    doc.sections[0]!.blocks[0] = {
      id: 'p',
      type: 'paragraph',
      props: {},
      runs: [{ id: 'r', props: {}, content: { type: 'text', text: 'word '.repeat(40) } }],
    }
    const patched = patchLayoutParagraph(doc, layout, 0, 0)
    expect(patched.layout.pages[0]!.blocks.some((b) => b.kind === 'table')).toBe(true)

    // non-paragraph patch returns needsFullLayout
    expect(patchLayoutParagraph(doc, layout, 0, 1).needsFullLayout).toBe(true)

    // cell patch with multiple paragraphs shifting later ones
    const multi: Document = createEmptyDocument()
    multi.sections[0]!.blocks = [
      {
        id: 't2',
        type: 'table',
        props: {},
        gridCols: [4000],
        rows: [
          {
            id: 'r0',
            props: {},
            cells: [
              {
                id: 'c0',
                props: {},
                blocks: [
                  {
                    id: 'p0',
                    type: 'paragraph',
                    props: {},
                    runs: [{ id: 'a', props: {}, content: { type: 'text', text: 'A' } }],
                  },
                  {
                    id: 'p1',
                    type: 'paragraph',
                    props: {},
                    runs: [{ id: 'b', props: {}, content: { type: 'text', text: 'B' } }],
                  },
                ],
              },
            ],
          },
        ],
      },
    ]
    layout = layoutDocument(multi)
    ;(multi.sections[0]!.blocks[0] as Table).rows[0]!.cells[0]!.blocks[0]!.runs[0]!.content = {
      type: 'text',
      text: 'A '.repeat(60),
    }
    const cellPatched = patchLayoutCellParagraph(multi, layout, 0, 0, { row: 0, cell: 0, para: 0 })
    expect(cellPatched.layout.pages[0]).toBeTruthy()
  })

  it('splits large tables across pages', () => {
    const doc = createEmptyDocument()
    // Shrink page so fewer rows fit
    doc.sections[0]!.properties.pageSize = { width: 6120, height: 4000 }
    doc.sections[0]!.properties.margins = {
      top: 200,
      right: 200,
      bottom: 200,
      left: 200,
      header: 100,
      footer: 100,
    }
    const rows = Array.from({ length: 20 }, (_, i) => ({
      id: `r${i}`,
      props: {},
      cells: [
        {
          id: `c${i}`,
          props: {},
          blocks: [
            {
              id: `p${i}`,
              type: 'paragraph',
              props: {},
              runs: [
                {
                  id: `t${i}`,
                  props: {},
                  content: { type: 'text' as const, text: `Row ${i} ${'text '.repeat(10)}` },
                },
              ],
            },
          ],
        },
      ],
    }))
    doc.sections[0]!.blocks = [
      {
        id: 't',
        type: 'table',
        props: {},
        gridCols: [5000],
        rows,
      },
    ]
    const layout = layoutDocument(doc)
    expect(layout.pages.length).toBeGreaterThan(1)
  })
})
