import {
  clampPosition,
  comparePositions,
  getParagraph,
  hitTestRun,
  normalizeRange,
  paragraphLength,
  paragraphPlainText,
  type DocPosition,
} from '../model/position.js'
import {
  coalesceRuns,
  createEmptyParagraph,
  deleteRangeInParagraph,
  insertTextInParagraph,
  setMarkInParagraph,
} from '../model/text.js'
import type {
  CharacterProps,
  Document,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from '../model/types.js'
import { assert } from '../util/assert.js'
import { nextId } from '../util/id.js'
import type { AppliedOp, EditorOp } from './types.js'

function updateParagraph(doc: Document, pos: DocPosition, paragraph: Paragraph): Document {
  const sections = doc.sections.map((section, si) => {
    if (si !== pos.sectionIndex) return section
    return {
      ...section,
      blocks: section.blocks.map((b, bi) => {
        if (bi !== pos.blockIndex) return b
        if (b.type === 'table' && pos.cell) {
          const { row, cell, para } = pos.cell
          return {
            ...b,
            rows: b.rows.map((r, ri) => {
              if (ri !== row) return r
              return {
                ...r,
                cells: r.cells.map((c, ci) => {
                  if (ci !== cell) return c
                  return {
                    ...c,
                    blocks: c.blocks.map((p, pi) => (pi === para ? paragraph : p)),
                  }
                }),
              }
            }),
          }
        }
        return paragraph
      }),
    }
  })
  return { ...doc, sections }
}

function captureDeleted(
  doc: Document,
  start: DocPosition,
  end: DocPosition,
): { text: string; props: CharacterProps } {
  // Phase 1: deletion within a single paragraph
  assert(
    start.sectionIndex === end.sectionIndex && start.blockIndex === end.blockIndex,
    'unsupported',
    'cross-paragraph delete not yet applied via this path',
  )
  const paragraph = getParagraph(doc, start)
  const text = paragraphPlainText(paragraph).slice(start.offset, end.offset)
  const hitProps = paragraph.runs[0]?.props ?? {}
  return { text, props: { ...hitProps } }
}

function getTable(doc: Document, sectionIndex: number, blockIndex: number): Table {
  const section = doc.sections[sectionIndex]
  assert(section, 'bad_position', 'section missing')
  const block = section.blocks[blockIndex]
  assert(block?.type === 'table', 'bad_position', 'expected table')
  return block
}

function replaceTable(
  doc: Document,
  sectionIndex: number,
  blockIndex: number,
  table: Table,
): Document {
  return {
    ...doc,
    sections: doc.sections.map((sec, i) =>
      i === sectionIndex
        ? { ...sec, blocks: sec.blocks.map((b, bi) => (bi === blockIndex ? table : b)) }
        : sec,
    ),
  }
}

/** Build an empty cell that mirrors the gridSpan/width of a template cell when given. */
function buildEmptyCell(template?: TableCell): TableCell {
  const props: TableCell['props'] = {}
  if (template?.props.gridSpan !== undefined) props.gridSpan = template.props.gridSpan
  if (template?.props.widthTwips !== undefined) props.widthTwips = template.props.widthTwips
  return {
    id: nextId('tc'),
    props,
    blocks: [createEmptyParagraph()],
  }
}

/** Build an empty row mirroring the gridSpan structure of a template row. */
function buildEmptyRow(template: TableRow): TableRow {
  return {
    id: nextId('tr'),
    props: {},
    cells: template.cells.map((c) => buildEmptyCell(c)),
  }
}

export function applyOp(doc: Document, op: EditorOp): { doc: Document; applied: AppliedOp } {
  switch (op.type) {
    case 'insertText': {
      const position = clampPosition(doc, op.position)
      const paragraph = getParagraph(doc, position)
      const next = insertTextInParagraph(paragraph, position.offset, op.text, op.props)
      const newDoc = updateParagraph(doc, position, next)
      const end: DocPosition = {
        ...position,
        offset: position.offset + op.text.length,
      }
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, position },
          inverse: {
            type: 'deleteRange',
            range: { anchor: position, focus: end },
            deletedText: op.text,
            ...(op.props !== undefined ? { deletedProps: op.props } : {}),
          },
        },
      }
    }
    case 'deleteRange': {
      const { start, end } = normalizeRange(op.range)
      const s = clampPosition(doc, start)
      const e = clampPosition(doc, end)
      if (comparePositions(s, e) === 0) {
        return {
          doc,
          applied: {
            forward: op,
            inverse: { type: 'insertText', position: s, text: '' },
          },
        }
      }

      if (s.sectionIndex === e.sectionIndex && s.blockIndex === e.blockIndex) {
        const captured = op.deletedText !== undefined ? null : captureDeleted(doc, s, e)
        const text = op.deletedText ?? captured!.text
        const props = op.deletedProps ?? captured?.props
        const paragraph = getParagraph(doc, s)
        const next = deleteRangeInParagraph(paragraph, s.offset, e.offset)
        return {
          doc: updateParagraph(doc, s, next),
          applied: {
            forward: {
              ...op,
              range: { anchor: s, focus: e },
              deletedText: text,
              ...(props !== undefined ? { deletedProps: props } : {}),
            },
            inverse: {
              type: 'insertText',
              position: s,
              text,
              ...(props !== undefined ? { props } : {}),
            },
          },
        }
      }

      // Multi-paragraph delete: join content
      assert(s.sectionIndex === e.sectionIndex, 'unsupported', 'cross-section delete')
      const section = doc.sections[s.sectionIndex]
      assert(section, 'bad_position', 'section missing')
      const first = getParagraph(doc, s)
      const last = getParagraph(doc, e)
      const head = paragraphPlainText(first).slice(0, s.offset)
      const tail = paragraphPlainText(last).slice(e.offset)
      const deletedParts: string[] = [paragraphPlainText(first).slice(s.offset)]
      for (let i = s.blockIndex + 1; i < e.blockIndex; i++) {
        const b = section.blocks[i]
        if (b?.type === 'paragraph') deletedParts.push(paragraphPlainText(b))
      }
      deletedParts.push(paragraphPlainText(last).slice(0, e.offset))
      const deletedText = op.deletedText ?? deletedParts.join('\n')

      const merged: Paragraph = {
        ...first,
        id: first.id,
        runs: [
          {
            id: nextId('r'),
            props: {},
            content: { type: 'text', text: head + tail },
          },
        ],
      }
      const blocks = [
        ...section.blocks.slice(0, s.blockIndex),
        merged,
        ...section.blocks.slice(e.blockIndex + 1),
      ]
      const newDoc: Document = {
        ...doc,
        sections: doc.sections.map((sec, i) =>
          i === s.sectionIndex ? { ...sec, blocks } : sec,
        ),
      }
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, deletedText, range: { anchor: s, focus: e } },
          inverse: {
            type: 'insertText',
            position: s,
            text: deletedText,
          },
        },
      }
    }
    case 'setMark': {
      const { start, end } = normalizeRange(op.range)
      const s = clampPosition(doc, start)
      const e = clampPosition(doc, end)
      assert(
        s.sectionIndex === e.sectionIndex && s.blockIndex === e.blockIndex,
        'unsupported',
        'setMark across paragraphs',
      )
      const paragraph = getParagraph(doc, s)
      const previousMarks = { ...op.mark }
      const next = setMarkInParagraph(paragraph, s.offset, e.offset, op.mark)
      // Invert: re-apply opposite boolean marks when possible
      const inverseMark: Partial<CharacterProps> = {}
      if (op.mark.bold !== undefined) inverseMark.bold = !op.mark.bold
      if (op.mark.italic !== undefined) inverseMark.italic = !op.mark.italic
      if (op.mark.underline !== undefined) inverseMark.underline = !op.mark.underline
      if (op.mark.strike !== undefined) inverseMark.strike = !op.mark.strike
      return {
        doc: updateParagraph(doc, s, next),
        applied: {
          forward: { ...op, range: { anchor: s, focus: e }, previousMarks },
          inverse: {
            type: 'setMark',
            range: { anchor: s, focus: e },
            mark: inverseMark,
          },
        },
      }
    }
    case 'splitParagraph': {
      const position = clampPosition(doc, op.position)
      const paragraph = getParagraph(doc, position)
      const plain = paragraphPlainText(paragraph)
      const leftPara: Paragraph = {
        ...paragraph,
        runs: [
          {
            id: nextId('r'),
            props: {},
            content: { type: 'text', text: plain.slice(0, position.offset) },
          },
        ],
      }
      const rightPara = createEmptyParagraph(paragraph.props.styleId ?? 'Normal')
      rightPara.runs = [
        {
          id: nextId('r'),
          props: {},
          content: { type: 'text', text: plain.slice(position.offset) },
        },
      ]
      rightPara.props = { ...paragraph.props }

      if (position.cell) {
        const section = doc.sections[position.sectionIndex]
        assert(section, 'bad_position', 'section missing')
        const table = section.blocks[position.blockIndex]
        assert(table?.type === 'table', 'bad_position', 'expected table')
        const { row, cell, para } = position.cell
        const newDoc = updateParagraph(doc, position, leftPara)
        const section2 = newDoc.sections[position.sectionIndex]!
        const table2 = section2.blocks[position.blockIndex]
        assert(table2?.type === 'table', 'bad_position', 'expected table')
        const nextTable = {
          ...table2,
          rows: table2.rows.map((r, ri) => {
            if (ri !== row) return r
            return {
              ...r,
              cells: r.cells.map((c, ci) => {
                if (ci !== cell) return c
                const blocks = [
                  ...c.blocks.slice(0, para),
                  leftPara,
                  rightPara,
                  ...c.blocks.slice(para + 1),
                ]
                return { ...c, blocks }
              }),
            }
          }),
        }
        const finalDoc: Document = {
          ...newDoc,
          sections: newDoc.sections.map((sec, i) =>
            i === position.sectionIndex
              ? {
                  ...sec,
                  blocks: sec.blocks.map((b, bi) => (bi === position.blockIndex ? nextTable : b)),
                }
              : sec,
          ),
        }
        return {
          doc: finalDoc,
          applied: {
            forward: { type: 'splitParagraph', position },
            inverse: {
              type: 'mergeParagraphs',
              position: {
                sectionIndex: position.sectionIndex,
                blockIndex: position.blockIndex,
                offset: 0,
                cell: { row, cell, para: para + 1 },
              },
              firstLength: position.offset,
            },
          },
        }
      }

      const section = doc.sections[position.sectionIndex]
      assert(section, 'bad_position', 'section missing')
      const blocks = [
        ...section.blocks.slice(0, position.blockIndex),
        leftPara,
        rightPara,
        ...section.blocks.slice(position.blockIndex + 1),
      ]
      const newDoc: Document = {
        ...doc,
        sections: doc.sections.map((sec, i) =>
          i === position.sectionIndex ? { ...sec, blocks } : sec,
        ),
      }
      return {
        doc: newDoc,
        applied: {
          forward: { type: 'splitParagraph', position },
          inverse: {
            type: 'mergeParagraphs',
            position: {
              sectionIndex: position.sectionIndex,
              blockIndex: position.blockIndex + 1,
              offset: 0,
            },
            firstLength: position.offset,
          },
        },
      }
    }
    case 'mergeParagraphs': {
      const position = clampPosition(doc, op.position)

      if (position.cell && position.cell.para > 0) {
        const firstPos: DocPosition = {
          sectionIndex: position.sectionIndex,
          blockIndex: position.blockIndex,
          offset: 0,
          cell: { ...position.cell, para: position.cell.para - 1 },
        }
        const first = getParagraph(doc, firstPos)
        const second = getParagraph(doc, position)
        const firstLen = op.firstLength ?? paragraphLength(first)
        const merged: Paragraph = {
          ...first,
          runs: [
            {
              id: nextId('r'),
              props: {},
              content: { type: 'text', text: paragraphPlainText(first) + paragraphPlainText(second) },
            },
          ],
        }
        const section = doc.sections[position.sectionIndex]
        assert(section, 'bad_position', 'section missing')
        const table = section.blocks[position.blockIndex]
        assert(table?.type === 'table', 'bad_position', 'expected table')
        const { row, cell, para } = position.cell
        const nextTable = {
          ...table,
          rows: table.rows.map((r, ri) => {
            if (ri !== row) return r
            return {
              ...r,
              cells: r.cells.map((c, ci) => {
                if (ci !== cell) return c
                const blocks = [
                  ...c.blocks.slice(0, para - 1),
                  merged,
                  ...c.blocks.slice(para + 1),
                ]
                return { ...c, blocks }
              }),
            }
          }),
        }
        const newDoc: Document = {
          ...doc,
          sections: doc.sections.map((sec, i) =>
            i === position.sectionIndex
              ? {
                  ...sec,
                  blocks: sec.blocks.map((b, bi) => (bi === position.blockIndex ? nextTable : b)),
                }
              : sec,
          ),
        }
        return {
          doc: newDoc,
          applied: {
            forward: { ...op, firstLength: firstLen },
            inverse: {
              type: 'splitParagraph',
              position: {
                sectionIndex: position.sectionIndex,
                blockIndex: position.blockIndex,
                offset: firstLen,
                cell: { row, cell, para: para - 1 },
              },
            },
          },
        }
      }

      assert(position.blockIndex > 0, 'bad_position', 'cannot merge first paragraph')
      const firstPos: DocPosition = {
        sectionIndex: position.sectionIndex,
        blockIndex: position.blockIndex - 1,
        offset: 0,
      }
      const first = getParagraph(doc, firstPos)
      const second = getParagraph(doc, position)
      const firstLen = op.firstLength ?? paragraphLength(first)
      const mergedText = paragraphPlainText(first) + paragraphPlainText(second)
      const merged: Paragraph = {
        ...first,
        runs: [
          {
            id: nextId('r'),
            props: {},
            content: { type: 'text', text: mergedText },
          },
        ],
      }
      const section = doc.sections[position.sectionIndex]
      assert(section, 'bad_position', 'section missing')
      const blocks = [
        ...section.blocks.slice(0, position.blockIndex - 1),
        merged,
        ...section.blocks.slice(position.blockIndex + 1),
      ]
      const newDoc: Document = {
        ...doc,
        sections: doc.sections.map((sec, i) =>
          i === position.sectionIndex ? { ...sec, blocks } : sec,
        ),
      }
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, firstLength: firstLen },
          inverse: {
            type: 'splitParagraph',
            position: {
              sectionIndex: position.sectionIndex,
              blockIndex: position.blockIndex - 1,
              offset: firstLen,
            },
          },
        },
      }
    }
    case 'setList': {
      const section = doc.sections[op.sectionIndex]
      assert(section, 'bad_position', 'section missing')
      const block = section.blocks[op.blockIndex]
      assert(block?.type === 'paragraph', 'bad_position', 'expected paragraph')
      const previous = block.props.numPr ?? null
      const nextProps =
        op.numPr === null
          ? (() => {
              const { numPr: _removed, ...rest } = block.props
              void _removed
              return rest
            })()
          : { ...block.props, numPr: op.numPr }
      const nextBlock = { ...block, props: nextProps }
      const newDoc: Document = {
        ...doc,
        sections: doc.sections.map((sec, i) =>
          i === op.sectionIndex
            ? {
                ...sec,
                blocks: sec.blocks.map((b, bi) => (bi === op.blockIndex ? nextBlock : b)),
              }
            : sec,
        ),
      }
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, previous },
          inverse: {
            type: 'setList',
            sectionIndex: op.sectionIndex,
            blockIndex: op.blockIndex,
            numPr: previous,
          },
        },
      }
    }
    case 'setParagraphProps': {
      const section = doc.sections[op.sectionIndex]
      assert(section, 'bad_position', 'section missing')
      const block = section.blocks[op.blockIndex]
      assert(block?.type === 'paragraph', 'bad_position', 'expected paragraph')
      const previous: Partial<typeof block.props> = {}
      for (const key of Object.keys(op.props) as (keyof typeof op.props)[]) {
        const v = block.props[key]
        if (v !== undefined) {
          ;(previous as Record<string, unknown>)[key as string] = v
        }
      }
      const nextBlock = { ...block, props: { ...block.props, ...op.props } }
      const newDoc: Document = {
        ...doc,
        sections: doc.sections.map((sec, i) =>
          i === op.sectionIndex
            ? {
                ...sec,
                blocks: sec.blocks.map((b, bi) => (bi === op.blockIndex ? nextBlock : b)),
              }
            : sec,
        ),
      }
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, previous },
          inverse: {
            type: 'setParagraphProps',
            sectionIndex: op.sectionIndex,
            blockIndex: op.blockIndex,
            props: previous,
          },
        },
      }
    }
    case 'insertTab': {
      const position = clampPosition(doc, op.position)
      const paragraph = getParagraph(doc, position)
      const hit = hitTestRun(paragraph, position.offset)
      const runs = [...paragraph.runs]
      const tabRun = {
        id: nextId('r'),
        props: { ...hit.run.props },
        content: { type: 'tab' as const },
      }
      if (hit.run.content.type === 'text' && hit.offsetInRun > 0) {
        const leftText = hit.run.content.text.slice(0, hit.offsetInRun)
        const rightText = hit.run.content.text.slice(hit.offsetInRun)
        runs.splice(
          hit.runIndex,
          1,
          { ...hit.run, id: nextId('r'), content: { type: 'text', text: leftText } },
          tabRun,
          {
            id: nextId('r'),
            props: { ...hit.run.props },
            content: { type: 'text', text: rightText },
          },
        )
      } else if (hit.offsetInRun === 0) {
        runs.splice(hit.runIndex, 0, tabRun)
      } else {
        runs.splice(hit.runIndex + 1, 0, tabRun)
      }
      const newDoc = updateParagraph(doc, position, {
        ...paragraph,
        runs: coalesceRuns(runs),
      })
      const end = { ...position, offset: position.offset + 1 }
      return {
        doc: newDoc,
        applied: {
          forward: { type: 'insertTab', position },
          inverse: {
            type: 'deleteRange',
            range: { anchor: position, focus: end },
            deletedText: '\t',
          },
        },
      }
    }
    case 'insertRow': {
      const table = getTable(doc, op.sectionIndex, op.blockIndex)
      assert(table.rows.length > 0, 'bad_position', 'table has no rows')
      const insertAt = Math.max(0, Math.min(op.afterRow + 1, table.rows.length))
      // Template row for gridSpan structure: prefer the row at afterRow, else the next row.
      const templateIndex =
        op.afterRow >= 0 ? Math.min(op.afterRow, table.rows.length - 1) : 0
      const template = table.rows[templateIndex]!
      const newRow = op.row ?? buildEmptyRow(template)
      const rows = [
        ...table.rows.slice(0, insertAt),
        newRow,
        ...table.rows.slice(insertAt),
      ]
      const newDoc = replaceTable(doc, op.sectionIndex, op.blockIndex, { ...table, rows })
      return {
        doc: newDoc,
        applied: {
          forward: { ...op },
          inverse: {
            type: 'deleteRow',
            sectionIndex: op.sectionIndex,
            blockIndex: op.blockIndex,
            row: insertAt,
          },
        },
      }
    }
    case 'deleteRow': {
      const table = getTable(doc, op.sectionIndex, op.blockIndex)
      assert(table.rows.length > 1, 'unsupported', 'cannot delete the only row in a table')
      assert(
        op.row >= 0 && op.row < table.rows.length,
        'bad_position',
        `table row ${String(op.row)} missing`,
      )
      const deletedRow = table.rows[op.row]!
      const rows = [...table.rows.slice(0, op.row), ...table.rows.slice(op.row + 1)]
      const newDoc = replaceTable(doc, op.sectionIndex, op.blockIndex, { ...table, rows })
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, deletedRow },
          inverse: {
            type: 'insertRow',
            sectionIndex: op.sectionIndex,
            blockIndex: op.blockIndex,
            afterRow: op.row - 1,
            row: deletedRow,
          },
        },
      }
    }
    case 'insertColumn': {
      const table = getTable(doc, op.sectionIndex, op.blockIndex)
      assert(table.rows.length > 0, 'bad_position', 'table has no rows')
      const insertAt = Math.max(0, op.afterCol + 1)
      const rows = table.rows.map((r, ri) => {
        const cellAt = Math.min(insertAt, r.cells.length)
        const newCell = op.cells?.[ri] ?? buildEmptyCell(r.cells[Math.min(op.afterCol, r.cells.length - 1)])
        return {
          ...r,
          cells: [...r.cells.slice(0, cellAt), newCell, ...r.cells.slice(cellAt)],
        }
      })
      let gridCols = table.gridCols
      if (gridCols) {
        const width =
          op.gridCol ??
          (gridCols.length > 0
            ? Math.round(gridCols.reduce((a, b) => a + b, 0) / gridCols.length)
            : 2400)
        const gridAt = Math.max(0, Math.min(insertAt, gridCols.length))
        gridCols = [...gridCols.slice(0, gridAt), width, ...gridCols.slice(gridAt)]
      }
      const newTable: Table = { ...table, rows }
      if (gridCols) newTable.gridCols = gridCols
      const newDoc = replaceTable(doc, op.sectionIndex, op.blockIndex, newTable)
      return {
        doc: newDoc,
        applied: {
          forward: { ...op },
          inverse: {
            type: 'deleteColumn',
            sectionIndex: op.sectionIndex,
            blockIndex: op.blockIndex,
            col: insertAt,
          },
        },
      }
    }
    case 'deleteColumn': {
      const table = getTable(doc, op.sectionIndex, op.blockIndex)
      const colCount = table.rows[0]?.cells.length ?? 0
      assert(colCount > 1, 'unsupported', 'cannot delete the only column in a table')
      assert(
        op.col >= 0 && op.col < colCount,
        'bad_position',
        `table column ${String(op.col)} missing`,
      )
      const deletedCells: TableCell[] = []
      const rows = table.rows.map((r) => {
        const removed = r.cells[op.col]
        if (removed) deletedCells.push(removed)
        return {
          ...r,
          cells: [...r.cells.slice(0, op.col), ...r.cells.slice(op.col + 1)],
        }
      })
      let gridCols = table.gridCols
      let deletedGridCol: number | undefined
      if (gridCols && op.col < gridCols.length) {
        deletedGridCol = gridCols[op.col]
        gridCols = [...gridCols.slice(0, op.col), ...gridCols.slice(op.col + 1)]
      }
      const newTable: Table = { ...table, rows }
      if (gridCols) newTable.gridCols = gridCols
      const newDoc = replaceTable(doc, op.sectionIndex, op.blockIndex, newTable)
      return {
        doc: newDoc,
        applied: {
          forward: { ...op, deletedCells, ...(deletedGridCol !== undefined ? { deletedGridCol } : {}) },
          inverse: {
            type: 'insertColumn',
            sectionIndex: op.sectionIndex,
            blockIndex: op.blockIndex,
            afterCol: op.col - 1,
            cells: deletedCells,
            ...(deletedGridCol !== undefined ? { gridCol: deletedGridCol } : {}),
          },
        },
      }
    }
  }
}
