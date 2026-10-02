import { describe, expect, it } from 'vitest'
import {
  applyOp,
  clampPosition,
  coalesceRuns,
  comparePositions,
  createEmptyDocument,
  createEmptyParagraph,
  createTextRun,
  deleteRangeInParagraph,
  documentCharCount,
  formatNumber,
  getParagraph,
  hitTestRun,
  IMAGE_PLACEHOLDER,
  insertTextInParagraph,
  isCollapsed,
  lookupLevel,
  nearestParagraphIndex,
  normalizeRange,
  paragraphLength,
  paragraphPlainText,
  positionsEqual,
  renderLevelText,
  resolveListMarkers,
  resolveParagraphProps,
  resolveRunProps,
  setMarkInParagraph,
  type Document,
  type NumberingLevel,
  type Run,
  type Table,
} from '../src/index.js'
import { getParagraphListMarker } from '../src/model/numbering.js'
import { inlineToChar } from '../src/model/position.js'
import { splitRun } from '../src/model/text.js'

describe('model numbering', () => {
  it('formats all number styles and resolves markers', () => {
    expect(formatNumber('decimal', 3)).toBe('3')
    expect(formatNumber('lowerLetter', 27)).toBe('aa')
    expect(formatNumber('upperLetter', 1)).toBe('A')
    expect(formatNumber('lowerRoman', 14)).toBe('xiv')
    expect(formatNumber('upperRoman', 4)).toBe('IV')
    expect(formatNumber('bullet', 1)).toBe('•')
    expect(formatNumber('none', 1)).toBe('')

    const levels: NumberingLevel[] = [
      { ilvl: 0, format: 'decimal', levelText: '%1.', start: 1 },
      { ilvl: 1, format: 'lowerLetter', levelText: '%1.%2)', start: 1 },
    ]
    expect(renderLevelText('%1.%2)', [1, 2], levels)).toBe('1.b)')
    expect(renderLevelText('%3.', [1], levels)).toBe('1.')

    let doc = createEmptyDocument()
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
      text: 'One',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 0,
      numPr: { numId: '2', ilvl: 0 },
    }).doc
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 0, offset: 3 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
      text: 'Nested',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 1,
      numPr: { numId: '2', ilvl: 1 },
    }).doc
    // custom abstract with font + startOverrides
    doc.numbering.abstractNums['abs_custom'] = {
      id: 'abs_custom',
      levels: [
        {
          ilvl: 0,
          format: 'upperRoman',
          levelText: '%1.',
          start: 2,
          indentLeft: 800,
          hanging: 400,
          fontFamily: 'Arial',
        },
      ],
    }
    doc.numbering.nums['9'] = {
      numId: '9',
      abstractNumId: 'abs_custom',
      startOverrides: { 0: 5 },
    }
    doc = applyOp(doc, {
      type: 'splitParagraph',
      position: { sectionIndex: 0, blockIndex: 1, offset: 6 },
    }).doc
    doc = applyOp(doc, {
      type: 'insertText',
      position: { sectionIndex: 0, blockIndex: 2, offset: 0 },
      text: 'Roman',
    }).doc
    doc = applyOp(doc, {
      type: 'setList',
      sectionIndex: 0,
      blockIndex: 2,
      numPr: { numId: '9', ilvl: 0 },
    }).doc

    const markers = resolveListMarkers(doc, 0)
    expect(markers.get(0)?.text).toMatch(/1/)
    expect(markers.get(2)?.fontFamily).toBe('Arial')
    expect(resolveListMarkers(doc, 99).size).toBe(0)
    expect(lookupLevel(doc.numbering, '2', 0)?.format).toBe('decimal')
    expect(lookupLevel(doc.numbering, 'missing', 0)).toBeUndefined()
    const para = doc.sections[0]!.blocks[0]!
    expect(getParagraphListMarker(doc, para as never, 0, 0)?.text).toBeTruthy()
    expect(
      getParagraphListMarker(
        doc,
        { ...(para as never), props: {} },
        0,
        0,
      ),
    ).toBeUndefined()
  })
})

describe('model position + text', () => {
  it('compares positions and clamps into tables', () => {
    const a = { sectionIndex: 0, blockIndex: 0, offset: 0 }
    const b = { sectionIndex: 0, blockIndex: 0, offset: 1 }
    expect(comparePositions(a, b)).toBeLessThan(0)
    expect(positionsEqual(a, a)).toBe(true)
    expect(isCollapsed({ anchor: a, focus: a })).toBe(true)
    expect(normalizeRange({ anchor: b, focus: a }).start).toEqual(a)

    const doc = createEmptyDocument()
    const table: Table = {
      id: 't',
      type: 'table',
      props: {},
      gridCols: [2000],
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
                  id: 'p',
                  type: 'paragraph',
                  props: {},
                  runs: [
                    { id: 'r1', props: {}, content: { type: 'text', text: 'Hi' } },
                    { id: 'r2', props: {}, content: { type: 'tab' } },
                    { id: 'r3', props: {}, content: { type: 'break', breakType: 'line' } },
                    {
                      id: 'r4',
                      props: {},
                      content: {
                        type: 'image',
                        mediaId: 'm1',
                        widthTwips: 100,
                        heightTwips: 100,
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    }
    doc.sections[0]!.blocks = [table]
    expect(inlineToChar(table.rows[0]!.cells[0]!.blocks[0]!.runs[3]!)).toBe(IMAGE_PLACEHOLDER)
    expect(paragraphPlainText(getParagraph(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }))).toContain('Hi')
    expect(clampPosition(doc, { sectionIndex: 0, blockIndex: 0, offset: 99 }).offset).toBeLessThan(99)
    expect(clampPosition(doc, { sectionIndex: 9, blockIndex: 9, offset: 0 }).sectionIndex).toBe(0)
    expect(nearestParagraphIndex(doc, 0, 0)).toBe(0)
    expect(documentCharCount(doc)).toBe(0) // table paras not counted at top level
    expect(paragraphLength(createEmptyParagraph())).toBe(0)

    const emptyDoc = createEmptyDocument()
    emptyDoc.sections[0]!.blocks = []
    expect(clampPosition(emptyDoc, { sectionIndex: 0, blockIndex: 0, offset: 0 })).toEqual({
      sectionIndex: 0,
      blockIndex: 0,
      offset: 0,
    })

    // compare cell paths
    expect(
      comparePositions(
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 1, para: 0 } },
      ),
    ).toBeLessThan(0)
    expect(
      comparePositions(
        { sectionIndex: 0, blockIndex: 0, offset: 0 },
        { sectionIndex: 0, blockIndex: 0, offset: 0, cell: { row: 0, cell: 0, para: 0 } },
      ),
    ).toBeLessThan(0)
  })

  it('splits/coalesces runs including non-text and marks images', () => {
    const textRun = createTextRun('abcd', { bold: true })
    const [l, r] = splitRun(textRun, 2)
    expect(l.content).toEqual({ type: 'text', text: 'ab' })
    expect(r.content).toEqual({ type: 'text', text: 'cd' })

    const img: Run = {
      id: 'img',
      props: {},
      content: { type: 'image', mediaId: 'm', widthTwips: 10, heightTwips: 10 },
    }
    expect(splitRun(img, 0)[0].content.type).toBe('text')
    expect(splitRun(img, 1)[1].content.type).toBe('text')

    const coalesced = coalesceRuns([
      createTextRun('a', { bold: true }),
      createTextRun('b', { bold: true }),
      createTextRun('', { bold: true }),
      createTextRun('c', { italic: true }),
      img,
    ])
    expect(coalesced[0]!.content).toEqual({ type: 'text', text: 'ab' })
    expect(coalesceRuns([])[0]!.content).toEqual({ type: 'text', text: '' })

    let para = createEmptyParagraph()
    para = insertTextInParagraph(para, 0, 'hello')
    para = insertTextInParagraph(para, 0, '')
    para = deleteRangeInParagraph(para, 1, 3)
    expect(paragraphPlainText(para)).toBe('hlo')
    expect(deleteRangeInParagraph(para, 2, 2)).toBe(para)

    // insert around image
    para = {
      ...para,
      runs: [img, createTextRun('z')],
    }
    para = insertTextInParagraph(para, 0, 'X')
    para = insertTextInParagraph(para, 2, 'Y')
    expect(paragraphPlainText(para)).toContain('X')

    para = setMarkInParagraph(
      {
        id: 'p',
        type: 'paragraph',
        props: {},
        runs: [
          createTextRun('ab', { bold: true }),
          img,
          { id: 't', props: {}, content: { type: 'tab' } },
          createTextRun('cd'),
        ],
      },
      0,
      4,
      { italic: true, bold: false },
    )
    expect(para.runs.some((r) => r.props.italic)).toBe(true)
    expect(setMarkInParagraph(para, 1, 1, { bold: true })).toBe(para)

    const hit = hitTestRun(createEmptyParagraph(), 0)
    expect(hit.offsetInRun).toBe(0)
  })
})

describe('styles cascade', () => {
  it('resolves paragraph and character style chains', () => {
    const doc = createEmptyDocument()
    doc.styles.paragraphStyles['Base'] = {
      id: 'Base',
      name: 'Base',
      basedOn: 'Normal',
      paragraph: { alignment: 'center', indentLeft: 100 },
      character: { fontSizePt: 14 },
    }
    doc.styles.paragraphStyles['Child'] = {
      id: 'Child',
      name: 'Child',
      basedOn: 'Base',
      paragraph: { alignment: 'right' },
      character: { bold: true },
    }
    doc.styles.characterStyles['Em'] = {
      id: 'Em',
      name: 'Em',
      basedOn: undefined,
      props: { italic: true, color: '#ff0000' },
    }
    doc.styles.characterStyles['Em2'] = {
      id: 'Em2',
      name: 'Em2',
      basedOn: 'Em',
      props: { underline: true },
    }
    // cycle should stop
    doc.styles.characterStyles['LoopA'] = {
      id: 'LoopA',
      name: 'LoopA',
      basedOn: 'LoopB',
      props: { strike: true },
    }
    doc.styles.characterStyles['LoopB'] = {
      id: 'LoopB',
      name: 'LoopB',
      basedOn: 'LoopA',
      props: { highlight: 'yellow' },
    }

    const para = {
      id: 'p',
      type: 'paragraph' as const,
      props: { styleId: 'Child', spacingAfter: 50 },
      runs: [
        {
          id: 'r',
          styleId: 'Em2',
          props: { color: '#00ff00' },
          content: { type: 'text' as const, text: 'x' },
        },
      ],
    }
    expect(resolveParagraphProps(doc, para).alignment).toBe('right')
    expect(resolveRunProps(doc, para, para.runs[0]!).italic).toBe(true)
    expect(resolveRunProps(doc, para, para.runs[0]!).underline).toBe(true)
    expect(resolveRunProps(doc, para, para.runs[0]!).color).toBe('#00ff00')

    const loopRun = { id: 'r2', styleId: 'LoopA', props: {}, content: { type: 'text' as const, text: 'y' } }
    expect(resolveRunProps(doc, para, loopRun).strike).toBe(true)

    // missing style id breaks chain
    const missing = {
      ...para,
      props: { styleId: 'Nope' },
    }
    expect(resolveParagraphProps(doc, missing).spacingAfter).toBe(200)
  })
})
