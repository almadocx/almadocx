import type { DocPosition, DocRange } from '../model/position.js'
import type { CharacterProps } from '../model/types.js'

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

export interface AppliedOp {
  forward: EditorOp
  inverse: EditorOp
}
