export { mountEditor, type EditorHandle, type EditorOptions, type ZoomMode } from './editor.js'
export { paintDocument, createCanvasMeasurer, measureContentSize } from './render/paint.js'
export { hitTestPoint, wordBounds, pageEditableParagraphs } from './selection/hitTest.js'
export { A11yMirror } from './a11y/mirror.js'
export { InputProxy } from './input/ime.js'
export {
  sanitizeHtml,
  htmlToFragment,
  plainTextToFragment,
  runsToHtml,
} from './clipboard/sanitize.js'
export { markdownToFragment } from './clipboard/markdown.js'
