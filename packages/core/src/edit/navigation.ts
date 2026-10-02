import type { Document } from '../model/types.js'
import {
  clampPosition,
  getParagraph,
  paragraphLength,
  paragraphPlainText,
  type CellPath,
  type DocPosition,
} from '../model/position.js'
import { moveByWord, paragraphBaseDirection } from './bidi.js'

function withCell(pos: DocPosition, cell: CellPath | undefined): DocPosition {
  if (!cell) {
    const { cell: _c, ...rest } = pos
    void _c
    return rest
  }
  return { ...pos, cell }
}

function firstCellInTable(doc: Document, sectionIndex: number, blockIndex: number): CellPath | undefined {
  const block = doc.sections[sectionIndex]?.blocks[blockIndex]
  if (!block || block.type !== 'table') return undefined
  for (let row = 0; row < block.rows.length; row++) {
    for (let cell = 0; cell < (block.rows[row]?.cells.length ?? 0); cell++) {
      if ((block.rows[row]!.cells[cell]?.blocks.length ?? 0) > 0) {
        return { row, cell, para: 0 }
      }
    }
  }
  return { row: 0, cell: 0, para: 0 }
}

function adjacentCell(
  doc: Document,
  pos: DocPosition,
  direction: -1 | 1,
): DocPosition | undefined {
  const block = doc.sections[pos.sectionIndex]?.blocks[pos.blockIndex]
  if (!block || block.type !== 'table' || !pos.cell) return undefined
  const { row, cell, para } = pos.cell
  const tc = block.rows[row]?.cells[cell]
  if (!tc) return undefined

  if (direction === 1) {
    if (para + 1 < tc.blocks.length) {
      return withCell({ ...pos, offset: 0 }, { row, cell, para: para + 1 })
    }
    // next cell in row
    if (cell + 1 < (block.rows[row]?.cells.length ?? 0)) {
      return withCell({ ...pos, offset: 0 }, { row, cell: cell + 1, para: 0 })
    }
    if (row + 1 < block.rows.length) {
      return withCell({ ...pos, offset: 0 }, { row: row + 1, cell: 0, para: 0 })
    }
    // exit table to next block
    const section = doc.sections[pos.sectionIndex]
    if (section && pos.blockIndex < section.blocks.length - 1) {
      const nextBi = pos.blockIndex + 1
      const nextBlock = section.blocks[nextBi]
      if (nextBlock?.type === 'table') {
        const cellPath = firstCellInTable(doc, pos.sectionIndex, nextBi)
        return withCell({ sectionIndex: pos.sectionIndex, blockIndex: nextBi, offset: 0 }, cellPath)
      }
      return { sectionIndex: pos.sectionIndex, blockIndex: nextBi, offset: 0 }
    }
    return undefined
  }

  // direction -1
  if (para > 0) {
    const prev = withCell({ ...pos, offset: 0 }, { row, cell, para: para - 1 })
    const p = getParagraph(doc, prev)
    return { ...prev, offset: paragraphLength(p) }
  }
  if (cell > 0) {
    const prevCell = block.rows[row]!.cells[cell - 1]!
    const paraIndex = Math.max(0, prevCell.blocks.length - 1)
    const prev = withCell({ ...pos, offset: 0 }, { row, cell: cell - 1, para: paraIndex })
    const p = getParagraph(doc, prev)
    return { ...prev, offset: paragraphLength(p) }
  }
  if (row > 0) {
    const prevRow = block.rows[row - 1]!
    const cellIndex = Math.max(0, prevRow.cells.length - 1)
    const prevCell = prevRow.cells[cellIndex]!
    const paraIndex = Math.max(0, prevCell.blocks.length - 1)
    const prev = withCell({ ...pos, offset: 0 }, { row: row - 1, cell: cellIndex, para: paraIndex })
    const p = getParagraph(doc, prev)
    return { ...prev, offset: paragraphLength(p) }
  }
  // exit table to previous block
  if (pos.blockIndex > 0) {
    const prevBi = pos.blockIndex - 1
    const prevBlock = doc.sections[pos.sectionIndex]?.blocks[prevBi]
    if (prevBlock?.type === 'table') {
      const cellPath = firstCellInTable(doc, pos.sectionIndex, prevBi)
      if (!cellPath) return { sectionIndex: pos.sectionIndex, blockIndex: prevBi, offset: 0 }
      // last cell of previous table
      const table = prevBlock
      let last: CellPath = cellPath
      for (let r = 0; r < table.rows.length; r++) {
        for (let c = 0; c < table.rows[r]!.cells.length; c++) {
          const n = table.rows[r]!.cells[c]!.blocks.length
          if (n > 0) last = { row: r, cell: c, para: n - 1 }
        }
      }
      const prev = withCell({ sectionIndex: pos.sectionIndex, blockIndex: prevBi, offset: 0 }, last)
      return { ...prev, offset: paragraphLength(getParagraph(doc, prev)) }
    }
    const prev = clampPosition(doc, { sectionIndex: pos.sectionIndex, blockIndex: prevBi, offset: 0 })
    return { ...prev, offset: paragraphLength(getParagraph(doc, prev)) }
  }
  return undefined
}

export function moveLeft(doc: Document, pos: DocPosition, byWord = false): DocPosition {
  const p = clampPosition(doc, pos)
  const para = getParagraph(doc, p)
  const text = paragraphPlainText(para)
  if (byWord) {
    const next = moveByWord(text, p.offset, -1)
    if (next !== p.offset || p.offset > 0) return { ...p, offset: next }
  } else if (p.offset > 0) {
    return { ...p, offset: p.offset - 1 }
  }
  const adj = adjacentCell(doc, p, -1)
  if (adj) return adj
  if (p.blockIndex > 0) {
    const prevPos = clampPosition(doc, {
      sectionIndex: p.sectionIndex,
      blockIndex: p.blockIndex - 1,
      offset: 0,
    })
    const prev = getParagraph(doc, prevPos)
    return { ...prevPos, offset: paragraphLength(prev) }
  }
  return p
}

export function moveRight(doc: Document, pos: DocPosition, byWord = false): DocPosition {
  const p = clampPosition(doc, pos)
  const para = getParagraph(doc, p)
  const text = paragraphPlainText(para)
  void paragraphBaseDirection(text)
  if (byWord) {
    const next = moveByWord(text, p.offset, 1)
    if (next !== p.offset || p.offset < text.length) return { ...p, offset: next }
  } else if (p.offset < text.length) {
    return { ...p, offset: p.offset + 1 }
  }
  const adj = adjacentCell(doc, p, 1)
  if (adj) return adj
  const section = doc.sections[p.sectionIndex]
  if (section && p.blockIndex < section.blocks.length - 1) {
    const nextBi = p.blockIndex + 1
    const nextBlock = section.blocks[nextBi]
    if (nextBlock?.type === 'table') {
      const cell = firstCellInTable(doc, p.sectionIndex, nextBi)
      return withCell({ sectionIndex: p.sectionIndex, blockIndex: nextBi, offset: 0 }, cell)
    }
    return clampPosition(doc, {
      sectionIndex: p.sectionIndex,
      blockIndex: nextBi,
      offset: 0,
    })
  }
  return p
}

export function moveHome(doc: Document, pos: DocPosition, toDocStart = false): DocPosition {
  const p = clampPosition(doc, pos)
  if (toDocStart) {
    const first = doc.sections[0]?.blocks[0]
    if (first?.type === 'table') {
      return withCell({ sectionIndex: 0, blockIndex: 0, offset: 0 }, firstCellInTable(doc, 0, 0))
    }
    return { sectionIndex: 0, blockIndex: 0, offset: 0 }
  }
  return { ...p, offset: 0 }
}

export function moveEnd(doc: Document, pos: DocPosition, toDocEnd = false): DocPosition {
  if (toDocEnd) {
    const lastSi = doc.sections.length - 1
    const section = doc.sections[lastSi]
    if (!section) return { sectionIndex: 0, blockIndex: 0, offset: 0 }
    let bi = section.blocks.length - 1
    const block = section.blocks[bi]
    if (block?.type === 'table') {
      let last: CellPath = { row: 0, cell: 0, para: 0 }
      for (let r = 0; r < block.rows.length; r++) {
        for (let c = 0; c < block.rows[r]!.cells.length; c++) {
          const n = block.rows[r]!.cells[c]!.blocks.length
          if (n > 0) last = { row: r, cell: c, para: n - 1 }
        }
      }
      const p = withCell({ sectionIndex: lastSi, blockIndex: bi, offset: 0 }, last)
      return { ...p, offset: paragraphLength(getParagraph(doc, p)) }
    }
    while (bi > 0 && section.blocks[bi]?.type !== 'paragraph') bi -= 1
    const para = getParagraph(doc, { sectionIndex: lastSi, blockIndex: bi, offset: 0 })
    return { sectionIndex: lastSi, blockIndex: bi, offset: paragraphLength(para) }
  }
  const p = clampPosition(doc, pos)
  const para = getParagraph(doc, p)
  return { ...p, offset: paragraphLength(para) }
}

/**
 * Word Tab / Shift+Tab: move to next/previous *cell* (not next paragraph in cell).
 * Returns needsNewRow when Tabbing from the last cell (caller inserts a row).
 */
export function moveTableCellTab(
  doc: Document,
  pos: DocPosition,
  direction: -1 | 1,
): { position: DocPosition; needsNewRow?: boolean } {
  const p = clampPosition(doc, pos)
  const block = doc.sections[p.sectionIndex]?.blocks[p.blockIndex]
  if (!block || block.type !== 'table' || !p.cell) {
    return { position: p }
  }
  const { row, cell } = p.cell
  if (direction === 1) {
    if (cell + 1 < (block.rows[row]?.cells.length ?? 0)) {
      return {
        position: withCell({ ...p, offset: 0 }, { row, cell: cell + 1, para: 0 }),
      }
    }
    if (row + 1 < block.rows.length) {
      return {
        position: withCell({ ...p, offset: 0 }, { row: row + 1, cell: 0, para: 0 }),
      }
    }
    return { position: p, needsNewRow: true }
  }
  // direction -1
  if (cell > 0) {
    const prevCell = block.rows[row]!.cells[cell - 1]!
    const paraIndex = Math.max(0, prevCell.blocks.length - 1)
    const next = withCell({ ...p, offset: 0 }, { row, cell: cell - 1, para: paraIndex })
    return { position: { ...next, offset: paragraphLength(getParagraph(doc, next)) } }
  }
  if (row > 0) {
    const prevRow = block.rows[row - 1]!
    const cellIndex = Math.max(0, prevRow.cells.length - 1)
    const prevCell = prevRow.cells[cellIndex]!
    const paraIndex = Math.max(0, prevCell.blocks.length - 1)
    const next = withCell({ ...p, offset: 0 }, { row: row - 1, cell: cellIndex, para: paraIndex })
    return { position: { ...next, offset: paragraphLength(getParagraph(doc, next)) } }
  }
  return { position: p }
}

/**
 * Prefer staying in the same table column (Word Up/Down).
 * Falls back to adjacentCell reading order when no same-column neighbor.
 */
export function moveTableCellVertical(
  doc: Document,
  pos: DocPosition,
  direction: -1 | 1,
  preferOffset?: number,
): DocPosition {
  const p = clampPosition(doc, pos)
  const block = doc.sections[p.sectionIndex]?.blocks[p.blockIndex]
  if (!block || block.type !== 'table' || !p.cell) return p

  const { row, cell, para } = p.cell
  const tc = block.rows[row]?.cells[cell]
  if (!tc) return p

  // Next/prev paragraph inside the same cell
  if (direction === 1 && para + 1 < tc.blocks.length) {
    const next = withCell({ ...p, offset: 0 }, { row, cell, para: para + 1 })
    const len = paragraphLength(getParagraph(doc, next))
    return { ...next, offset: Math.max(0, Math.min(preferOffset ?? p.offset, len)) }
  }
  if (direction === -1 && para > 0) {
    const next = withCell({ ...p, offset: 0 }, { row, cell, para: para - 1 })
    const len = paragraphLength(getParagraph(doc, next))
    return { ...next, offset: Math.max(0, Math.min(preferOffset ?? p.offset, len)) }
  }

  // Same column, adjacent row
  const targetRow = row + direction
  if (targetRow >= 0 && targetRow < block.rows.length) {
    const rowCells = block.rows[targetRow]!.cells
    const ci = Math.max(0, Math.min(cell, rowCells.length - 1))
    const targetCell = rowCells[ci]!
    if (targetCell.props.vMerge === 'continue') {
      // Skip continue cells — walk further in direction
      return moveTableCellVertical(
        doc,
        withCell({ ...p, offset: 0 }, { row: targetRow, cell: ci, para: 0 }),
        direction,
        preferOffset,
      )
    }
    const paraIndex = direction === 1 ? 0 : Math.max(0, targetCell.blocks.length - 1)
    const next = withCell({ ...p, offset: 0 }, { row: targetRow, cell: ci, para: paraIndex })
    const len = paragraphLength(getParagraph(doc, next))
    return { ...next, offset: Math.max(0, Math.min(preferOffset ?? p.offset, len)) }
  }

  const adj = adjacentCell(doc, p, direction)
  if (adj) {
    const len = paragraphLength(getParagraph(doc, adj))
    return { ...adj, offset: Math.max(0, Math.min(preferOffset ?? p.offset, len)) }
  }
  return p
}

/**
 * Vertical movement using layout line metrics when available.
 * Fallback: jump ±1 paragraph / cell.
 */
export function moveVertical(
  doc: Document,
  pos: DocPosition,
  direction: -1 | 1,
  preferOffset?: number,
): DocPosition {
  const p = clampPosition(doc, pos)
  if (p.cell) {
    return moveTableCellVertical(doc, p, direction, preferOffset)
  }
  const section = doc.sections[p.sectionIndex]
  if (!section) return p
  let bi = p.blockIndex + direction
  while (bi >= 0 && bi < section.blocks.length) {
    const block = section.blocks[bi]
    if (block?.type === 'paragraph') {
      const para = getParagraph(doc, { sectionIndex: p.sectionIndex, blockIndex: bi, offset: 0 })
      const len = paragraphLength(para)
      const offset = Math.max(0, Math.min(preferOffset ?? p.offset, len))
      return { sectionIndex: p.sectionIndex, blockIndex: bi, offset }
    }
    if (block?.type === 'table') {
      const cell = firstCellInTable(doc, p.sectionIndex, bi)
      const next = withCell({ sectionIndex: p.sectionIndex, blockIndex: bi, offset: 0 }, cell)
      const para = getParagraph(doc, next)
      const len = paragraphLength(para)
      return { ...next, offset: Math.max(0, Math.min(preferOffset ?? p.offset, len)) }
    }
    bi += direction
  }
  return p
}
