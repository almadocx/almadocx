import type { DocPosition, DocRange } from '../model/position.js'
import type { CharacterProps, TableCell, TableRow } from '../model/types.js'

export type { DocPosition, DocRange }

export type EditorOp =
  | {
      type: 'insertText'
      position: DocPosition
      text: string
      props?: CharacterProps
    }
  | {
      type: 'deleteRange'
      range: DocRange
      deletedText?: string
      deletedProps?: CharacterProps
    }
  | {
      type: 'setMark'
      range: DocRange
      mark: Partial<CharacterProps>
      previousMarks?: Partial<CharacterProps>
    }
  | {
      type: 'splitParagraph'
      position: DocPosition
    }
  | {
      type: 'mergeParagraphs'
      position: DocPosition
      firstLength?: number
    }
  | {
      type: 'setList'
      blockIndex: number
      sectionIndex: number
      /** null clears list */
      numPr: { numId: string; ilvl: number } | null
      previous?: { numId: string; ilvl: number } | null
    }
  | {
      type: 'setParagraphProps'
      sectionIndex: number
      blockIndex: number
      props: Partial<import('../model/types.js').ParagraphProps>
      previous?: Partial<import('../model/types.js').ParagraphProps>
    }
  | {
      type: 'insertTab'
      position: DocPosition
    }
  | {
      type: 'insertRow'
      sectionIndex: number
      blockIndex: number
      /** Insert the new row after this row index; use -1 to insert before the first row. */
      afterRow: number
      /**
       * Exact row to insert. When omitted, an empty row is generated that mirrors the
       * gridSpan structure of the template row. Used by `deleteRow`'s inverse to restore
       * the original row verbatim.
       */
      row?: TableRow
    }
  | {
      type: 'deleteRow'
      sectionIndex: number
      blockIndex: number
      row: number
      /** Captured row for the inverse. */
      deletedRow?: TableRow
    }
  | {
      type: 'insertColumn'
      sectionIndex: number
      blockIndex: number
      /** Insert the new cell after this cell index in every row; use -1 to insert before the first. */
      afterCol: number
      /**
       * Exact cells to insert (one per row). When omitted, empty cells are generated.
       * Used by `deleteColumn`'s inverse to restore the original column verbatim.
       */
      cells?: TableCell[]
      /** Exact grid column width to restore (used by `deleteColumn`'s inverse). */
      gridCol?: number
    }
  | {
      type: 'deleteColumn'
      sectionIndex: number
      blockIndex: number
      col: number
      /** Captured cells (one per row) for the inverse. */
      deletedCells?: TableCell[]
      /** Captured grid column width for the inverse. */
      deletedGridCol?: number
    }

export interface AppliedOp {
  forward: EditorOp
  inverse: EditorOp
}
