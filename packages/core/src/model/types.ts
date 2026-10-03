/**
 * Format-neutral document model for Almadocx.
 * Layout and render must never import OOXML/ODF types — only this model.
 */

export type DocumentId = string
export type BlockId = string
export type RunId = string
export type StyleId = string
export type MediaId = string
export type NumId = string
export type AbstractNumId = string

export type LengthTwips = number

export interface CharacterProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  fontFamily?: string
  fontSizePt?: number
  color?: string
  highlight?: string
  verticalAlign?: 'baseline' | 'superscript' | 'subscript'
}

export type NumberFormat =
  | 'decimal'
  | 'lowerLetter'
  | 'upperLetter'
  | 'lowerRoman'
  | 'upperRoman'
  | 'bullet'
  | 'none'

export interface NumberingLevel {
  ilvl: number
  format: NumberFormat
  /** Pattern with %1..%9 placeholders, e.g. "%1." or "•" */
  levelText: string
  start: number
  indentLeft?: LengthTwips
  hanging?: LengthTwips
  alignment?: 'left' | 'center' | 'right'
  fontFamily?: string
}

export interface AbstractNumbering {
  id: AbstractNumId
  levels: NumberingLevel[]
}

export interface NumberingInstance {
  numId: NumId
  abstractNumId: AbstractNumId
  /** Per-level start overrides */
  startOverrides?: Record<number, number>
}

export interface NumberingDefinitions {
  abstractNums: Record<AbstractNumId, AbstractNumbering>
  nums: Record<NumId, NumberingInstance>
}

export interface NumPr {
  numId: NumId
  ilvl: number
}

export interface TabStop {
  position: LengthTwips
  alignment?: 'left' | 'center' | 'right' | 'decimal'
  leader?: 'none' | 'dot' | 'dash' | 'underscore'
}

export interface ParagraphProps {
  alignment?: 'left' | 'center' | 'right' | 'justify'
  indentLeft?: LengthTwips
  indentRight?: LengthTwips
  indentFirstLine?: LengthTwips
  spacingBefore?: LengthTwips
  spacingAfter?: LengthTwips
  lineSpacing?: number
  lineSpacingRule?: 'auto' | 'exact' | 'atLeast'
  styleId?: StyleId
  numPr?: NumPr
  keepNext?: boolean
  keepLines?: boolean
  widowControl?: boolean
  pageBreakBefore?: boolean
  /** Explicit tab stops; default every 720 twips (0.5") when empty */
  tabs?: TabStop[]
}

export type InlineContent =
  | { type: 'text'; text: string }
  | { type: 'tab' }
  | { type: 'break'; breakType: 'line' | 'page' | 'column' }
  | {
      type: 'image'
      mediaId: MediaId
      widthTwips: LengthTwips
      heightTwips: LengthTwips
      alt?: string
    }

export interface Run {
  id: RunId
  props: CharacterProps
  styleId?: StyleId
  content: InlineContent
}

export interface Paragraph {
  id: BlockId
  type: 'paragraph'
  props: ParagraphProps
  runs: Run[]
}

export interface TableBorders {
  top?: string
  bottom?: string
  left?: string
  right?: string
  insideH?: string
  insideV?: string
}

export interface TableProps {
  widthTwips?: LengthTwips
  alignment?: 'left' | 'center' | 'right'
  borders?: TableBorders
  cellSpacing?: LengthTwips
  /** Table left indent (tblInd) in twips */
  indentTwips?: LengthTwips
}

export interface CellMargin {
  top?: LengthTwips
  right?: LengthTwips
  bottom?: LengthTwips
  left?: LengthTwips
}

export interface CellBorders {
  top?: string
  bottom?: string
  left?: string
  right?: string
}

export interface TableCellProps {
  widthTwips?: LengthTwips
  shading?: string
  vAlign?: 'top' | 'center' | 'bottom'
  /** Grid span for horizontal merge */
  gridSpan?: number
  /** Vertical merge: restart or continue */
  vMerge?: 'restart' | 'continue'
  /** Cell margins (tcMar) in twips */
  margin?: CellMargin
  /** Cell border colors (tcBorders) */
  borders?: CellBorders
}

export interface TableCell {
  id: string
  props: TableCellProps
  blocks: Paragraph[]
}

export interface TableRowProps {
  heightTwips?: LengthTwips
  header?: boolean
}

export interface TableRow {
  id: string
  props: TableRowProps
  cells: TableCell[]
}

export interface Table {
  id: BlockId
  type: 'table'
  props: TableProps
  rows: TableRow[]
  /** Column widths in twips when known */
  gridCols?: LengthTwips[]
}

export type Block = Paragraph | Table

export interface PageMargins {
  top: LengthTwips
  right: LengthTwips
  bottom: LengthTwips
  left: LengthTwips
  header: LengthTwips
  footer: LengthTwips
}

export interface PageSize {
  width: LengthTwips
  height: LengthTwips
}

export interface HeaderFooter {
  id: string
  blocks: Paragraph[]
}

export interface SectionProperties {
  pageSize: PageSize
  margins: PageMargins
  columns: number
  titlePage?: boolean
}

export interface Section {
  id: string
  properties: SectionProperties
  blocks: Block[]
  header?: HeaderFooter
  footer?: HeaderFooter
  headerFirst?: HeaderFooter
  footerFirst?: HeaderFooter
}

export interface CharacterStyle {
  id: StyleId
  name: string
  basedOn?: StyleId
  props: CharacterProps
}

export interface ParagraphStyle {
  id: StyleId
  name: string
  basedOn?: StyleId
  paragraph: ParagraphProps
  character: CharacterProps
}

export interface StyleSheet {
  docDefaults: {
    paragraph: ParagraphProps
    character: CharacterProps
  }
  /** Style marked `w:default="1"` (usually Normal); used when a paragraph omits pStyle. */
  defaultParagraphStyle?: StyleId
  paragraphStyles: Record<StyleId, ParagraphStyle>
  characterStyles: Record<StyleId, CharacterStyle>
}

export interface MediaItem {
  id: MediaId
  contentType: string
  bytes: Uint8Array
  /** Original package path when known */
  path?: string
}

/** Opaque package parts preserved for lossless same-format round-trip. */
export interface PreservedPart {
  path: string
  bytes: Uint8Array
  contentType?: string
}

export type SourceFormat = 'docx' | 'odt' | 'internal'

export interface DocumentPackageMeta {
  sourceFormat: SourceFormat
  preservedParts: PreservedPart[]
  macroPartPaths: string[]
  contentTypesXml?: string
  relationships?: Record<string, string>
}

export interface Document {
  id: DocumentId
  styles: StyleSheet
  numbering: NumberingDefinitions
  media: Record<MediaId, MediaItem>
  sections: Section[]
  package: DocumentPackageMeta
}

export const DEFAULT_PAGE_SIZE: PageSize = {
  width: 12240,
  height: 15840,
}

export const DEFAULT_MARGINS: PageMargins = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
}

export function emptyNumbering(): NumberingDefinitions {
  return { abstractNums: {}, nums: {} }
}

/** Built-in bullet and decimal lists for editing without a numbering part. */
export function ensureDefaultNumbering(defs: NumberingDefinitions): NumberingDefinitions {
  const next = {
    abstractNums: { ...defs.abstractNums },
    nums: { ...defs.nums },
  }
  if (!next.abstractNums['abs_bullet']) {
    next.abstractNums['abs_bullet'] = {
      id: 'abs_bullet',
      levels: [
        {
          ilvl: 0,
          format: 'bullet',
          levelText: '•',
          start: 1,
          indentLeft: 720,
          hanging: 360,
        },
        {
          ilvl: 1,
          format: 'bullet',
          levelText: '○',
          start: 1,
          indentLeft: 1440,
          hanging: 360,
        },
      ],
    }
  }
  if (!next.abstractNums['abs_decimal']) {
    next.abstractNums['abs_decimal'] = {
      id: 'abs_decimal',
      levels: [
        {
          ilvl: 0,
          format: 'decimal',
          levelText: '%1.',
          start: 1,
          indentLeft: 720,
          hanging: 360,
        },
        {
          ilvl: 1,
          format: 'decimal',
          levelText: '%1.%2.',
          start: 1,
          indentLeft: 1440,
          hanging: 360,
        },
      ],
    }
  }
  if (!next.nums['1']) {
    next.nums['1'] = { numId: '1', abstractNumId: 'abs_bullet' }
  }
  if (!next.nums['2']) {
    next.nums['2'] = { numId: '2', abstractNumId: 'abs_decimal' }
  }
  return next
}

export function createEmptyDocument(format: SourceFormat = 'internal'): Document {
  return {
    id: 'doc_1',
    styles: {
      docDefaults: {
        paragraph: {
          spacingAfter: 200,
          lineSpacing: 1.15,
          lineSpacingRule: 'auto',
          widowControl: true,
        },
        character: {
          fontFamily: 'Liberation Serif',
          fontSizePt: 11,
          color: '#000000',
        },
      },
      paragraphStyles: {
        Normal: {
          id: 'Normal',
          name: 'Normal',
          paragraph: {},
          character: {},
        },
      },
      characterStyles: {},
    },
    numbering: ensureDefaultNumbering(emptyNumbering()),
    media: {},
    sections: [
      {
        id: 'sect_1',
        properties: {
          pageSize: { ...DEFAULT_PAGE_SIZE },
          margins: { ...DEFAULT_MARGINS },
          columns: 1,
        },
        blocks: [
          {
            id: 'p_1',
            type: 'paragraph',
            props: { styleId: 'Normal' },
            runs: [
              {
                id: 'r_1',
                props: {},
                content: { type: 'text', text: '' },
              },
            ],
          },
        ],
      },
    ],
    package: {
      sourceFormat: format,
      preservedParts: [],
      macroPartPaths: [],
    },
  }
}
