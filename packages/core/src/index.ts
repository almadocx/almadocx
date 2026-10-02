export { AlmadocxError, assert } from './util/assert.js'
export {
  TWIPS_PER_INCH,
  inchesToTwips,
  pointsToTwips,
  twipsToPx,
  pxToTwips,
} from './util/units.js'
export { nextId, resetIdsForTests } from './util/id.js'

export type {
  Document,
  Section,
  Paragraph,
  Run,
  Block,
  Table,
  TableCell,
  TableRow,
  CharacterProps,
  ParagraphProps,
  StyleSheet,
  SourceFormat,
  DocumentPackageMeta,
  PreservedPart,
  NumberingDefinitions,
  NumPr,
  MediaItem,
  HeaderFooter,
} from './model/types.js'
export {
  createEmptyDocument,
  DEFAULT_MARGINS,
  DEFAULT_PAGE_SIZE,
  emptyNumbering,
  ensureDefaultNumbering,
} from './model/types.js'
export {
  type DocPosition,
  type DocRange,
  type CellPath,
  comparePositions,
  positionsEqual,
  normalizeRange,
  isCollapsed,
  paragraphPlainText,
  paragraphLength,
  getParagraph,
  clampPosition,
  documentCharCount,
  hitTestRun,
  IMAGE_PLACEHOLDER,
  nearestParagraphIndex,
} from './model/position.js'
export {
  createTextRun,
  createEmptyParagraph,
  insertTextInParagraph,
  deleteRangeInParagraph,
  setMarkInParagraph,
  coalesceRuns,
} from './model/text.js'
export {
  resolveListMarkers,
  formatNumber,
  renderLevelText,
  lookupLevel,
  type ResolvedListMarker,
} from './model/numbering.js'
export { FONT_SUBSTITUTES, resolveFontFamily, fontStackCss } from './fonts/substitutes.js'

export { resolveParagraphProps, resolveRunProps } from './styles/cascade.js'

export type { EditorOp, AppliedOp } from './ops/types.js'
export { applyOp } from './ops/apply.js'
export {
  type HistoryState,
  createHistory,
  dispatch,
  dispatchBatch,
  undo,
  redo,
  canUndo,
  canRedo,
} from './ops/history.js'

export {
  findAll,
  findNext,
  replaceMatch,
  replaceAll,
  extractPlainRange,
  type FindOptions,
  type FindMatch,
} from './edit/findReplace.js'
export {
  charDirection,
  paragraphBaseDirection,
  moveCaretByArrow,
  moveByWord,
  isRtlChar,
} from './edit/bidi.js'
export {
  moveLeft,
  moveRight,
  moveHome,
  moveEnd,
  moveVertical,
  moveTableCellTab,
  moveTableCellVertical,
} from './edit/navigation.js'

export { nextTabStopTwips, defaultTabs } from './layout/tabs.js'

export { PACKAGE_LIMITS } from './io/limits.js'
export { readZip, writeZip, isMacroPath } from './io/zip.js'
export { parseXml, buildXml } from './io/xml.js'

export { loadDocument, saveDocument, detectFormat } from './formats/load.js'
export { parseDocx, sniffDocx } from './formats/docx/parse.js'
export { serializeDocx } from './formats/docx/serialize.js'
export { parseOdt, sniffOdt } from './formats/odt/parse.js'
export { serializeOdt } from './formats/odt/serialize.js'

export type {
  LayoutResult,
  LayoutPage,
  LayoutParagraph,
  LayoutTable,
  LayoutTableCell,
  LayoutBlock,
  LayoutLine,
  LayoutGlyphRun,
  LayoutOptions,
  PatchLayoutResult,
} from './layout/layout.js'
export { layoutDocument, patchLayoutParagraph, patchLayoutCellParagraph } from './layout/layout.js'
export {
  type TextMeasurer,
  type TextMetrics,
  createApproximateMeasurer,
  fontCss,
} from './layout/measure.js'
