import {
  canRedo,
  canUndo,
  createEmptyDocument,
  createHistory,
  dispatch,
  dispatchBatch,
  extractPlainRange,
  findAll,
  findNext,
  getParagraph,
  hitTestRun,
  layoutDocument,
  loadDocument,
  moveEnd,
  moveHome,
  moveLeft,
  moveRight,
  moveTableCellTab,
  moveVertical,
  normalizeRange,
  isCollapsed,
  paragraphLength,
  paragraphPlainText,
  patchLayoutParagraph,
  patchLayoutCellParagraph,
  redo,
  replaceMatch,
  saveDocument,
  undo,
  type DocPosition,
  type DocRange,
  type Document,
  type EditorOp,
  type FindMatch,
  type FindOptions,
  type HistoryState,
  type LayoutResult,
  type SourceFormat,
  type CharacterProps,
} from '@almadocx/core'
import { A11yMirror } from './a11y/mirror.js'
import { markdownToFragment } from './clipboard/markdown.js'
import { htmlToFragment, plainTextToFragment, runsToHtml, sanitizeHtml } from './clipboard/sanitize.js'
import { InputProxy } from './input/ime.js'
import { createCanvasMeasurer, getCaretScreenRect, measureContentSize, paintDocument } from './render/paint.js'
import { PAGE_GAP_DEFAULT, PAGE_TOP_PAD } from './render/pageLayout.js'
import { hitTestPoint, wordBounds } from './selection/hitTest.js'
import { moveByLayoutLine, offsetToX } from './selection/selectionPaint.js'
import { pageEditableParagraphs } from './selection/hitTest.js'

export type ZoomMode = 'percent' | 'fitWidth' | 'fitPage' | 'twoPages'

export interface EditorOptions {
  document?: Document
  zoom?: number
  zoomMode?: ZoomMode
  onChange?: (doc: Document) => void
}

export interface EditorHandle {
  getDocument(): Document
  setDocument(doc: Document): void
  loadBytes(bytes: Uint8Array, format?: SourceFormat): void
  saveBytes(format?: SourceFormat): Uint8Array
  focus(): void
  destroy(): void
  getSelection(): DocRange
  undo(): void
  redo(): void
  setZoom(zoom: number): void
  getZoom(): number
  setZoomMode(mode: ZoomMode): void
  getZoomMode(): ZoomMode
  bold(): void
  italic(): void
  underline(): void
  strike(): void
  align(alignment: 'left' | 'center' | 'right' | 'justify'): void
  bulletList(): void
  numberedList(): void
  clearList(): void
  find(options: FindOptions): FindMatch[]
  findNext(options: FindOptions): FindMatch | undefined
  replace(options: FindOptions, replacement: string): boolean
  replaceAll(options: FindOptions, replacement: string): number
  selectAll(): void
}

export function mountEditor(host: HTMLElement, options: EditorOptions = {}): EditorHandle {
  host.classList.add('almadocx-host')
  Object.assign(host.style, {
    position: 'relative',
    overflow: 'auto',
    outline: 'none',
    width: '100%',
    height: '100%',
    minHeight: '0',
  })
  host.tabIndex = 0

  const canvas = document.createElement('canvas')
  canvas.className = 'almadocx-canvas'
  const spacer = document.createElement('div')
  spacer.className = 'almadocx-spacer'
  spacer.style.position = 'relative'
  spacer.style.width = '100%'
  spacer.style.pointerEvents = 'none'
  spacer.appendChild(canvas)
  host.appendChild(spacer)
  // Canvas receives pointer events
  canvas.style.pointerEvents = 'auto'
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Canvas 2D unavailable')

  let doc = options.document ?? createEmptyDocument()
  let history: HistoryState = createHistory()
  let selection: DocRange = {
    anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
    focus: { sectionIndex: 0, blockIndex: 0, offset: 0 },
  }
  let zoom = options.zoom ?? 1
  let zoomMode: ZoomMode = options.zoomMode ?? 'percent'
  let pageColumns = 1
  const measurer = createCanvasMeasurer(ctx)
  let layout: LayoutResult = layoutDocument(doc, { measurer })
  let caretBlink = true
  let composingPreview: string | undefined
  let dragging = false
  let dragMoving = false
  let dragSource: DocRange | undefined
  // Stable anchor for the active pointer drag (survives rect-selection anchor rewrites).
  let dragAnchor: DocPosition | undefined
  // True while dragging a Word-style rectangular multi-cell table selection.
  let rectSelecting = false
  let preferredCol: number | undefined
  let preferredX: number | undefined
  let pendingMarks: CharacterProps = {}
  let selectAllLevel = 0 // 0=none, 1=cell, 2=table, 3=doc (Ctrl+A escalation)
  let lastClickAt = 0
  let lastClickPos: DocPosition | undefined
  let clickCount = 0
  let fullLayoutTimer: ReturnType<typeof setTimeout> | null = null
  let notifyTimer: ReturnType<typeof setTimeout> | null = null
  const imageCache = new Map<string, HTMLImageElement>()
  const objectUrls: string[] = []

  const a11y = new A11yMirror(host)

  const cloneRange = (range: DocRange): DocRange => ({
    anchor: {
      ...range.anchor,
      ...(range.anchor.cell ? { cell: { ...range.anchor.cell } } : {}),
    },
    focus: {
      ...range.focus,
      ...(range.focus.cell ? { cell: { ...range.focus.cell } } : {}),
    },
  })

  const loadImages = () => {
    for (const url of objectUrls) URL.revokeObjectURL(url)
    objectUrls.length = 0
    imageCache.clear()
    for (const [id, media] of Object.entries(doc.media)) {
      if (!media.bytes.byteLength) continue
      if (!media.contentType.startsWith('image/')) continue
      const blob = new Blob([new Uint8Array(media.bytes)], { type: media.contentType })
      const url = URL.createObjectURL(blob)
      objectUrls.push(url)
      const img = new Image()
      img.decoding = 'async'
      img.onload = () => {
        imageCache.set(id, img)
        paintOnly()
      }
      img.src = url
    }
  }

  const resolveZoom = () => {
    const page = layout.pages[0]
    if (!page) return
    const availW = Math.max(1, host.clientWidth)
    const availH = Math.max(1, host.clientHeight - PAGE_TOP_PAD * 2)
    if (zoomMode === 'fitWidth') {
      pageColumns = 1
      // Exact fit: page edge-to-edge with viewport — no leftover H-scroll
      zoom = Math.max(0.25, Math.min(4, availW / page.width))
    } else if (zoomMode === 'fitPage') {
      pageColumns = 1
      zoom = Math.max(0.25, Math.min(4, Math.min(availW / page.width, availH / page.height)))
    } else if (zoomMode === 'twoPages') {
      pageColumns = 2
      zoom = Math.max(0.25, Math.min(4, availW / (page.width * 2 + PAGE_GAP_DEFAULT)))
    } else {
      pageColumns = 1
    }
  }

  const scrollCaretIntoView = () => {
    const rect = getCaretScreenRect(layout, selection, zoom, PAGE_GAP_DEFAULT, pageColumns, host.clientWidth)
    if (!rect) return
    const pad = 48
    const viewTop = host.scrollTop
    const viewBottom = viewTop + host.clientHeight
    if (rect.top < viewTop + pad) {
      host.scrollTop = Math.max(0, rect.top - pad)
    } else if (rect.bottom > viewBottom - pad) {
      host.scrollTop = rect.bottom - host.clientHeight + pad
    }
    // Only scroll horizontally when content is genuinely wider than the viewport
    if (zoomMode !== 'fitWidth' && zoomMode !== 'fitPage') {
      const contentW = measureContentSize(layout, zoom, PAGE_GAP_DEFAULT, pageColumns).width
      if (contentW > host.clientWidth + 1) {
        const viewLeft = host.scrollLeft
        const viewRight = viewLeft + host.clientWidth
        if (rect.left < viewLeft + pad) {
          host.scrollLeft = Math.max(0, rect.left - pad)
        } else if (rect.left > viewRight - pad) {
          host.scrollLeft = rect.left - host.clientWidth + pad
        }
      } else {
        host.scrollLeft = 0
      }
    } else {
      host.scrollLeft = 0
    }
  }

  const paintOnly = () => {
    resolveZoom()
    // Stabilize width: always reserve vertical scrollbar gutter so centering doesn't shift
    host.style.overflowY = 'scroll'
    const vw = host.clientWidth
    const size = measureContentSize(layout, zoom, PAGE_GAP_DEFAULT, pageColumns)
    // Spacer fills the host horizontally so the sticky canvas covers the full view;
    // when zoomed past viewport width, minWidth enables H-scroll.
    spacer.style.width = '100%'
    spacer.style.minWidth = `${Math.ceil(size.width)}px`
    spacer.style.height = `${Math.ceil(size.height)}px`
    const fitsWidth = size.width <= vw + 0.5
    const noHScroll = zoomMode === 'fitWidth' || zoomMode === 'fitPage' || fitsWidth
    host.style.overflowX = noHScroll ? 'hidden' : 'auto'
    if (noHScroll) host.scrollLeft = 0
    paintDocument(canvas, ctx, {
      layout,
      doc,
      selection,
      zoom,
      dpr: window.devicePixelRatio || 1,
      caretVisible: (caretBlink || Boolean(composingPreview)) && isCollapsed(selection),
      images: imageCache,
      virtualized: true,
      pageColumns,
      viewport: {
        scrollTop: host.scrollTop,
        scrollLeft: host.scrollLeft,
        width: vw,
        height: host.clientHeight,
      },
    })
  }

  const scheduleNotify = () => {
    if (notifyTimer) clearTimeout(notifyTimer)
    notifyTimer = setTimeout(() => {
      notifyTimer = null
      a11y.sync(doc)
      options.onChange?.(doc)
    }, 100)
  }

  const scheduleFullLayout = (delayMs = 48) => {
    if (fullLayoutTimer) clearTimeout(fullLayoutTimer)
    fullLayoutTimer = setTimeout(() => {
      fullLayoutTimer = null
      layout = layoutDocument(doc, { measurer })
      paintOnly()
      scrollCaretIntoView()
    }, delayMs)
  }

  /** Full document reflow — structural edits, undo, paste, formatting. */
  const render = () => {
    if (fullLayoutTimer) {
      clearTimeout(fullLayoutTimer)
      fullLayoutTimer = null
    }
    layout = layoutDocument(doc, { measurer })
    paintOnly()
    scrollCaretIntoView()
    scheduleNotify()
  }

  /** Selection / caret move — paint only, no layout. */
  const renderSelection = () => {
    paintOnly()
    scrollCaretIntoView()
  }

  /**
   * Typing / single-char delete: patch one paragraph in place, paint immediately,
   * and debounce a full reflow for page breaks / widows.
   */
  const renderTyping = (pos: DocPosition) => {
    if (pos.cell) {
      const patched = patchLayoutCellParagraph(
        doc,
        layout,
        pos.sectionIndex,
        pos.blockIndex,
        pos.cell,
        { measurer },
      )
      layout = patched.layout
      paintOnly()
      scrollCaretIntoView()
      scheduleNotify()
      scheduleFullLayout(patched.needsFullLayout ? 48 : 160)
      return
    }
    const patched = patchLayoutParagraph(doc, layout, pos.sectionIndex, pos.blockIndex, { measurer })
    layout = patched.layout
    paintOnly()
    scrollCaretIntoView()
    scheduleNotify()
    scheduleFullLayout(patched.needsFullLayout ? 48 : 160)
  }

  const commitOp = (op: EditorOp, after: DocRange, mode: 'full' | 'fast' = 'full') => {
    const before = cloneRange(selection)
    ;({ doc, history } = dispatch(doc, history, op, { before, after: cloneRange(after) }))
    selection = after
    if (mode === 'fast') renderTyping(after.focus)
    else render()
  }

  const commitOps = (ops: EditorOp[], after: DocRange, mode: 'full' | 'fast' = 'full') => {
    if (ops.length === 0) return
    const before = cloneRange(selection)
    if (ops.length === 1) {
      ;({ doc, history } = dispatch(doc, history, ops[0]!, { before, after: cloneRange(after) }))
    } else {
      ;({ doc, history } = dispatchBatch(doc, history, ops, { before, after: cloneRange(after) }))
    }
    selection = after
    if (mode === 'fast') renderTyping(after.focus)
    else render()
  }

  const input = new InputProxy(host, {
    onCommit: (text) => {
      if (text === '\n') {
        commitSplit()
        return
      }
      insertText(text)
    },
    onCompositionUpdate: (text) => {
      composingPreview = text
      paintOnly()
    },
    onCompositionEnd: (text) => {
      composingPreview = undefined
      if (text) insertText(text)
      else paintOnly()
    },
  })
  // keydown wired after onKeyDown is defined

  /** Character style to the left of the caret (Word/Docs typing inheritance). */
  const typingStyleAt = (pos: DocPosition): CharacterProps => {
    const para = getParagraph(doc, pos)
    if (para.runs.length === 0) return {}
    if (pos.offset <= 0) {
      return { ...para.runs[0]!.props }
    }
    const hit = hitTestRun(para, pos.offset - 1)
    return { ...hit.run.props }
  }

  /**
   * Merge left-of-caret style with pending toggles.
   * `pendingMarks[key] === false` forces that mark off for subsequent typing.
   */
  const resolveInsertProps = (pos: DocPosition, extra?: CharacterProps): CharacterProps => {
    const out: CharacterProps = { ...typingStyleAt(pos), ...pendingMarks, ...(extra ?? {}) }
    for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
      if (out[key] === false) delete out[key]
    }
    return out
  }

  const clearPendingMarks = () => {
    if (Object.keys(pendingMarks).length > 0) pendingMarks = {}
  }

  const insertText = (text: string, props?: CharacterProps) => {
    // Typing over a rectangular multi-cell selection clears every selected cell
    // then inserts the text in the top-left cell (Word behavior).
    const rectBounds = rectSelecting ? rectSelectionBounds() : undefined
    if (rectBounds) {
      const cleared = buildRectClearOps(rectBounds)
      if (cleared) {
        const mergedProps = resolveInsertProps(cleared.caret, props)
        const ops: EditorOp[] = [
          ...cleared.ops,
          { type: 'insertText', position: cleared.caret, text, props: mergedProps },
        ]
        const focus: DocPosition = { ...cleared.caret, offset: cleared.caret.offset + text.length }
        rectSelecting = false
        caretBlink = true
        preferredCol = focus.offset
        commitOps(ops, { anchor: focus, focus }, 'fast')
        return
      }
    }
    const { start } = normalizeRange(selection)
    const ops: EditorOp[] = []
    let caret = selection.focus
    if (!isCollapsed(selection)) {
      ops.push({ type: 'deleteRange', range: selection })
      caret = start
    }
    const mergedProps = resolveInsertProps(caret, props)
    ops.push({
      type: 'insertText',
      position: caret,
      text,
      props: mergedProps,
    })
    const focus: DocPosition = { ...caret, offset: caret.offset + text.length }
    const after: DocRange = { anchor: focus, focus }
    caretBlink = true
    preferredCol = focus.offset
    // Pending toggles stay until the caret moves (Word sticky formatting)
    commitOps(ops, after, 'fast')
  }

  const insertTab = () => {
    const { start } = normalizeRange(selection)
    const ops: EditorOp[] = []
    let caret = selection.focus
    if (!isCollapsed(selection)) {
      ops.push({ type: 'deleteRange', range: selection })
      caret = start
    }
    ops.push({ type: 'insertTab', position: caret })
    const focus = { ...caret, offset: caret.offset + 1 }
    commitOps(ops, { anchor: focus, focus }, 'fast')
  }

  const commitSplit = () => {
    const { start, end } = normalizeRange(selection)
    const ops: EditorOp[] = []
    let caret = selection.focus
    if (start.offset !== end.offset || start.blockIndex !== end.blockIndex) {
      ops.push({ type: 'deleteRange', range: selection })
      caret = start
    }
    const prevBlock = doc.sections[caret.sectionIndex]?.blocks[caret.blockIndex]
    const prevNum = prevBlock?.type === 'paragraph' ? prevBlock.props.numPr : undefined
    const isEmptyListItem =
      prevBlock?.type === 'paragraph' &&
      Boolean(prevNum) &&
      paragraphPlainText(prevBlock).length === 0 &&
      caret.offset === 0

    // Word/Docs: Enter on empty list item exits the list
    if (isEmptyListItem) {
      ops.push({
        type: 'setList',
        sectionIndex: caret.sectionIndex,
        blockIndex: caret.blockIndex,
        numPr: null,
      })
      preferredCol = 0
      commitOps(ops, { anchor: caret, focus: caret })
      return
    }

    ops.push({ type: 'splitParagraph', position: caret })
    let next: DocPosition
    if (caret.cell) {
      next = {
        sectionIndex: caret.sectionIndex,
        blockIndex: caret.blockIndex,
        offset: 0,
        cell: { ...caret.cell, para: caret.cell.para + 1 },
      }
    } else {
      next = {
        sectionIndex: caret.sectionIndex,
        blockIndex: caret.blockIndex + 1,
        offset: 0,
      }
      if (prevNum) {
        ops.push({
          type: 'setList',
          sectionIndex: next.sectionIndex,
          blockIndex: next.blockIndex,
          numPr: { ...prevNum },
        })
      }
    }
    preferredCol = 0
    commitOps(ops, { anchor: next, focus: next })
  }

  /** Last paragraph index of a given table cell (0 when empty/missing). */
  const lastCellPara = (
    sectionIndex: number,
    blockIndex: number,
    row: number,
    cell: number,
  ): number => {
    const block = doc.sections[sectionIndex]?.blocks[blockIndex]
    if (block?.type !== 'table') return 0
    const tc = block.rows[row]?.cells[cell]
    return Math.max(0, (tc?.blocks.length ?? 1) - 1)
  }

  /**
   * Rectangular table-cell selection bounds, when the current selection spans
   * multiple cells of a single table block. Returns undefined otherwise.
   */
  const rectSelectionBounds = ():
    | { sectionIndex: number; blockIndex: number; r0: number; c0: number; r1: number; c1: number }
    | undefined => {
    const { start, end } = normalizeRange(selection)
    if (!start.cell || !end.cell) return undefined
    if (start.sectionIndex !== end.sectionIndex || start.blockIndex !== end.blockIndex) {
      return undefined
    }
    if (start.cell.row === end.cell.row && start.cell.cell === end.cell.cell) return undefined
    return {
      sectionIndex: start.sectionIndex,
      blockIndex: start.blockIndex,
      r0: Math.min(start.cell.row, end.cell.row),
      c0: Math.min(start.cell.cell, end.cell.cell),
      r1: Math.max(start.cell.row, end.cell.row),
      c1: Math.max(start.cell.cell, end.cell.cell),
    }
  }

  /**
   * Build ops that clear the text of every cell in the rectangle, plus the
   * caret position (top-left cell, first paragraph). deleteRange is applied
   * per cell paragraph since it cannot cross cells.
   */
  const buildRectClearOps = (bounds: {
    sectionIndex: number
    blockIndex: number
    r0: number
    c0: number
    r1: number
    c1: number
  }): { ops: EditorOp[]; caret: DocPosition } | undefined => {
    const { sectionIndex, blockIndex, r0, c0, r1, c1 } = bounds
    const table = doc.sections[sectionIndex]?.blocks[blockIndex]
    if (table?.type !== 'table') return undefined
    const ops: EditorOp[] = []
    for (let row = r0; row <= r1; row++) {
      const tableRow = table.rows[row]
      if (!tableRow) continue
      for (let cell = c0; cell <= c1; cell++) {
        const tc = tableRow.cells[cell]
        if (!tc) continue
        for (let para = 0; para < tc.blocks.length; para++) {
          const block = tc.blocks[para]
          if (block?.type !== 'paragraph') continue
          const len = paragraphLength(block)
          if (len === 0) continue
          const from: DocPosition = {
            sectionIndex,
            blockIndex,
            offset: 0,
            cell: { row, cell, para },
          }
          const to: DocPosition = {
            sectionIndex,
            blockIndex,
            offset: len,
            cell: { row, cell, para },
          }
          ops.push({ type: 'deleteRange', range: { anchor: from, focus: to } })
        }
      }
    }
    const caret: DocPosition = {
      sectionIndex,
      blockIndex,
      offset: 0,
      cell: { row: r0, cell: c0, para: 0 },
    }
    return { ops, caret }
  }

  /**
   * When a rectangular multi-cell selection is active, clear every selected
   * cell's content and collapse the caret to the top-left cell. Returns true
   * when it handled the delete.
   */
  const clearRectSelection = (): boolean => {
    if (!rectSelecting) return false
    const bounds = rectSelectionBounds()
    if (!bounds) {
      rectSelecting = false
      return false
    }
    const cleared = buildRectClearOps(bounds)
    rectSelecting = false
    if (!cleared) return false
    const after: DocRange = { anchor: cleared.caret, focus: cleared.caret }
    preferredCol = 0
    if (cleared.ops.length === 0) {
      selection = after
      renderSelection()
      return true
    }
    commitOps(cleared.ops, after)
    return true
  }

  const deleteBackward = () => {
    if (clearRectSelection()) return
    const { start, end } = normalizeRange(selection)
    if (
      start.offset !== end.offset ||
      start.blockIndex !== end.blockIndex ||
      start.sectionIndex !== end.sectionIndex
    ) {
      commitOp({ type: 'deleteRange', range: selection }, { anchor: start, focus: start }, 'fast')
      return
    }
    if (start.offset > 0) {
      const from = { ...start, offset: start.offset - 1 }
      preferredCol = from.offset
      commitOp(
        { type: 'deleteRange', range: { anchor: from, focus: start } },
        { anchor: from, focus: from },
        'fast',
      )
      return
    }
    if (start.cell && start.cell.para > 0) {
      const prevPara = getParagraph(doc, {
        ...start,
        cell: { ...start.cell, para: start.cell.para - 1 },
        offset: 0,
      })
      const firstLen = paragraphLength(prevPara)
      const after: DocPosition = {
        ...start,
        offset: firstLen,
        cell: { ...start.cell, para: start.cell.para - 1 },
      }
      preferredCol = firstLen
      commitOp({ type: 'mergeParagraphs', position: start }, { anchor: after, focus: after })
      return
    }
    // At start of first para in a cell — move to previous cell (Word-like), don't delete table
    if (start.cell && start.cell.para === 0 && start.offset === 0) {
      const { position } = moveTableCellTab(doc, start, -1)
      if (
        position.cell &&
        (position.cell.row !== start.cell.row || position.cell.cell !== start.cell.cell)
      ) {
        setCaret(position, false)
      }
      return
    }
    if (start.blockIndex > 0 && !start.cell) {
      const prev = getParagraph(doc, {
        sectionIndex: start.sectionIndex,
        blockIndex: start.blockIndex - 1,
        offset: 0,
      })
      const firstLen = paragraphLength(prev)
      const after: DocPosition = {
        sectionIndex: start.sectionIndex,
        blockIndex: start.blockIndex - 1,
        offset: firstLen,
      }
      preferredCol = firstLen
      commitOp({ type: 'mergeParagraphs', position: start }, { anchor: after, focus: after })
    }
  }

  const deleteForward = () => {
    if (clearRectSelection()) return
    const { start } = normalizeRange(selection)
    if (!isCollapsed(selection)) {
      commitOp({ type: 'deleteRange', range: selection }, { anchor: start, focus: start }, 'fast')
      return
    }
    const next = moveRight(doc, start)
    if (
      next.sectionIndex === start.sectionIndex &&
      next.blockIndex === start.blockIndex &&
      next.offset === start.offset
    ) {
      return
    }
    commitOp(
      { type: 'deleteRange', range: { anchor: start, focus: next } },
      { anchor: start, focus: start },
      'fast',
    )
  }

  const selectedParagraphBlocks = (): { sectionIndex: number; blockIndex: number }[] => {
    const { start, end } = normalizeRange(selection)
    const out: { sectionIndex: number; blockIndex: number }[] = []
    for (let si = start.sectionIndex; si <= end.sectionIndex; si++) {
      const section = doc.sections[si]
      if (!section) continue
      const from = si === start.sectionIndex ? start.blockIndex : 0
      const to = si === end.sectionIndex ? end.blockIndex : section.blocks.length - 1
      for (let bi = from; bi <= to; bi++) {
        if (section.blocks[bi]?.type === 'paragraph') out.push({ sectionIndex: si, blockIndex: bi })
      }
    }
    return out.length ? out : [{ sectionIndex: start.sectionIndex, blockIndex: start.blockIndex }]
  }

  const applyList = (numPr: { numId: string; ilvl: number } | null) => {
    const targets = selectedParagraphBlocks()
    // Toggle off when every selected paragraph already has this list
    let clear = numPr === null
    if (numPr) {
      clear = targets.every((t) => {
        const b = doc.sections[t.sectionIndex]?.blocks[t.blockIndex]
        return (
          b?.type === 'paragraph' &&
          b.props.numPr?.numId === numPr.numId &&
          b.props.numPr?.ilvl === numPr.ilvl
        )
      })
    }
    const effective = clear ? null : numPr
    const ops: EditorOp[] = targets.map((t) => ({
      type: 'setList',
      sectionIndex: t.sectionIndex,
      blockIndex: t.blockIndex,
      numPr: effective,
    }))
    commitOps(ops, cloneRange(selection))
    input.focus()
  }

  /** True when every character in the selection already has this boolean mark. */
  const selectionHasBooleanMark = (key: 'bold' | 'italic' | 'underline' | 'strike'): boolean => {
    const { start, end } = normalizeRange(selection)
    if (
      start.sectionIndex !== end.sectionIndex ||
      start.blockIndex !== end.blockIndex ||
      start.offset === end.offset
    ) {
      // Multi-para: require all paragraphs' selected spans (simplify: single-para for now)
      if (start.sectionIndex !== end.sectionIndex || start.blockIndex !== end.blockIndex) {
        // Walk each para in range
        for (let bi = start.blockIndex; bi <= end.blockIndex; bi++) {
          const block = doc.sections[start.sectionIndex]?.blocks[bi]
          if (block?.type !== 'paragraph') continue
          const from = bi === start.blockIndex ? start.offset : 0
          const to = bi === end.blockIndex ? end.offset : paragraphLength(block)
          if (to <= from) continue
          for (let o = from; o < to; o++) {
            const hit = hitTestRun(block, o)
            if (!hit.run.props[key]) return false
          }
        }
        return true
      }
    }
    const para = getParagraph(doc, start)
    if (end.offset <= start.offset) return false
    for (let o = start.offset; o < end.offset; o++) {
      const hit = hitTestRun(para, o)
      if (!hit.run.props[key]) return false
    }
    return true
  }

  const applyMark = (mark: Partial<CharacterProps>) => {
    const { start, end } = normalizeRange(selection)
    const collapsed =
      start.sectionIndex === end.sectionIndex &&
      start.blockIndex === end.blockIndex &&
      start.offset === end.offset &&
      ((!start.cell && !end.cell) ||
        (start.cell &&
          end.cell &&
          start.cell.row === end.cell.row &&
          start.cell.cell === end.cell.cell &&
          start.cell.para === end.cell.para))
    if (collapsed) {
      // Toggle sticky overrides against the current typing style (inherited + pending)
      const current = resolveInsertProps(selection.focus)
      const next: CharacterProps = { ...pendingMarks }
      for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
        if (mark[key] !== true) continue
        if (current[key]) {
          // Force off even when adjacent text is bold/italic
          ;(next as Record<string, unknown>)[key] = false
        } else {
          next[key] = true
        }
      }
      for (const [k, v] of Object.entries(mark) as [keyof CharacterProps, CharacterProps[keyof CharacterProps]][]) {
        if (k === 'bold' || k === 'italic' || k === 'underline' || k === 'strike') continue
        if (v !== undefined) (next as Record<string, unknown>)[k as string] = v
      }
      pendingMarks = next
      input.focus()
      return
    }
    // Range formatting applies only to the selection — do not sticky-persist
    const toggled: Partial<CharacterProps> = { ...mark }
    for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
      if (mark[key] === true && selectionHasBooleanMark(key)) {
        toggled[key] = false
      }
    }
    commitOp(
      {
        type: 'setMark',
        range: selection,
        mark: toggled,
      },
      cloneRange(selection),
    )
    pendingMarks = {}
    input.focus()
  }

  const applyAlign = (alignment: 'left' | 'center' | 'right' | 'justify') => {
    const targets = selectedParagraphBlocks()
    const ops: EditorOp[] = targets.map((t) => ({
      type: 'setParagraphProps',
      sectionIndex: t.sectionIndex,
      blockIndex: t.blockIndex,
      props: { alignment },
    }))
    commitOps(ops, cloneRange(selection))
    input.focus()
  }

  const pasteFragment = (html: string | undefined, plain: string) => {
    let fragment =
      html && html.trim()
        ? htmlToFragment(sanitizeHtml(html))
        : markdownToFragment(plain) ?? plainTextToFragment(plain)
    // Prefer markdown structure when HTML is a single dumped paragraph of MD-looking text
    if (html && fragment.paragraphs.length <= 1) {
      const md = markdownToFragment(plain)
      if (md && md.paragraphs.length > 1) fragment = md
    }
    const before = cloneRange(selection)
    const ops: EditorOp[] = []
    const { start } = normalizeRange(selection)
    let caret = selection.focus
    if (!isCollapsed(selection)) {
      ops.push({ type: 'deleteRange', range: selection })
      caret = start
    }
    for (let i = 0; i < fragment.paragraphs.length; i++) {
      const p = fragment.paragraphs[i]!
      if (i > 0) {
        ops.push({ type: 'splitParagraph', position: caret })
        if (caret.cell) {
          caret = {
            ...caret,
            cell: { ...caret.cell, para: caret.cell.para + 1 },
            offset: 0,
          }
        } else {
          caret = { sectionIndex: caret.sectionIndex, blockIndex: caret.blockIndex + 1, offset: 0 }
        }
      }
      if (p.props.numPr) {
        ops.push({
          type: 'setList',
          sectionIndex: caret.sectionIndex,
          blockIndex: caret.blockIndex,
          numPr: p.props.numPr,
        })
      }
      if (p.props.styleId && p.props.styleId !== 'Normal') {
        ops.push({
          type: 'setParagraphProps',
          sectionIndex: caret.sectionIndex,
          blockIndex: caret.blockIndex,
          props: { styleId: p.props.styleId },
        })
      }
      for (const run of p.runs) {
        if (run.content.type === 'text' && run.content.text) {
          ops.push({
            type: 'insertText',
            position: caret,
            text: run.content.text,
            props: { ...run.props },
          })
          caret = { ...caret, offset: caret.offset + run.content.text.length }
        } else if (run.content.type === 'tab') {
          ops.push({ type: 'insertTab', position: caret })
          caret = { ...caret, offset: caret.offset + 1 }
        }
      }
    }
    const after: DocRange = { anchor: caret, focus: caret }
    if (ops.length === 1) {
      ;({ doc, history } = dispatch(doc, history, ops[0]!, { before, after }))
    } else if (ops.length > 1) {
      ;({ doc, history } = dispatchBatch(doc, history, ops, { before, after }))
    }
    selection = after
    preferredCol = caret.offset
    render()
  }

  const copySelection = async (cut: boolean) => {
    const { start, end } = normalizeRange(selection)
    if (isCollapsed(selection)) return
    const plain = extractPlainRange(doc, { anchor: start, focus: end })
    // Single-paragraph HTML for Phase 3
    const block = doc.sections[start.sectionIndex]?.blocks[start.blockIndex]
    let html = `<p>${plain.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</p>`
    if (block?.type === 'paragraph' && start.blockIndex === end.blockIndex) {
      html = runsToHtml([
        {
          ...block,
          runs: block.runs, // full para — acceptable for MVP rich copy within para marks
        },
      ])
    }
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/plain': new Blob([plain], { type: 'text/plain' }),
          'text/html': new Blob([html], { type: 'text/html' }),
        }),
      ])
    } catch {
      await navigator.clipboard.writeText(plain)
    }
    if (cut) {
      commitOp({ type: 'deleteRange', range: selection }, { anchor: start, focus: start })
    }
  }

  const onPaste = async (e: ClipboardEvent) => {
    e.preventDefault()
    const html = e.clipboardData?.getData('text/html')
    const plain = e.clipboardData?.getData('text/plain') ?? ''
    pasteFragment(html || undefined, plain)
  }

  const onCopy = (e: ClipboardEvent) => {
    e.preventDefault()
    void copySelection(false)
  }

  const onCut = (e: ClipboardEvent) => {
    e.preventDefault()
    void copySelection(true)
  }

  const caretPreferX = (pos: DocPosition): number | undefined => {
    for (const page of layout.pages) {
      for (const para of pageEditableParagraphs(page)) {
        if (para.sectionIndex !== pos.sectionIndex || para.blockIndex !== pos.blockIndex) continue
        if (pos.cell || para.cell) {
          if (
            !pos.cell ||
            !para.cell ||
            pos.cell.row !== para.cell.row ||
            pos.cell.cell !== para.cell.cell ||
            pos.cell.para !== para.cell.para
          ) {
            continue
          }
        }
        for (let li = 0; li < para.lines.length; li++) {
          const line = para.lines[li]!
          const isLast = li === para.lines.length - 1
          const inLine =
            pos.offset >= line.startOffset &&
            (pos.offset < line.endOffset || (pos.offset === line.endOffset && isLast))
          if (!inLine) continue
          return offsetToX(line, pos.offset)
        }
      }
    }
    return undefined
  }

  const setCaret = (pos: DocPosition, extend: boolean) => {
    rectSelecting = false
    if (!extend) {
      clearPendingMarks()
      selectAllLevel = 0
    }
    selection = extend ? { anchor: selection.anchor, focus: pos } : { anchor: pos, focus: pos }
    preferredCol = pos.offset
    preferredX = caretPreferX(pos)
    caretBlink = true
    renderSelection()
  }

  const moveCaretVertical = (direction: -1 | 1, extend: boolean) => {
    const focus = selection.focus
    const x = preferredX ?? caretPreferX(focus)
    const byLine = moveByLayoutLine(layout, focus, direction, x)
    if (byLine) {
      preferredX = x
      setCaret(byLine, extend)
      preferredX = x // setCaret overwrites; restore sticky X
      preferredCol = byLine.offset
      return
    }
    setCaret(moveVertical(doc, focus, direction, preferredCol), extend)
    if (x !== undefined) preferredX = x
  }

  const applyHistory = (dir: 'undo' | 'redo') => {
    const result = dir === 'undo' ? undo(doc, history) : redo(doc, history)
    doc = result.doc
    history = result.history
    if (result.selection) selection = result.selection
    clearPendingMarks()
    selectAllLevel = 0
    caretBlink = true
    preferredCol = selection.focus.offset
    preferredX = caretPreferX(selection.focus)
    render()
  }

  const onKeyDown = (e: KeyboardEvent) => {
    if (input.isComposing()) return
    const mod = e.metaKey || e.ctrlKey
    if (mod && e.key.toLowerCase() === 'z' && !e.shiftKey) {
      e.preventDefault()
      if (canUndo(history)) applyHistory('undo')
      return
    }
    if (mod && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) {
      e.preventDefault()
      if (canRedo(history)) applyHistory('redo')
      return
    }
    if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault()
      selectAll()
      return
    }
    if (mod && e.key.toLowerCase() === 'b') {
      e.preventDefault()
      applyMark({ bold: true })
      return
    }
    if (mod && e.key.toLowerCase() === 'i') {
      e.preventDefault()
      applyMark({ italic: true })
      return
    }
    if (mod && e.key.toLowerCase() === 'u') {
      e.preventDefault()
      applyMark({ underline: true })
      return
    }
    if (mod && e.key.toLowerCase() === 'c') {
      e.preventDefault()
      void copySelection(false)
      return
    }
    if (mod && e.key.toLowerCase() === 'x') {
      e.preventDefault()
      void copySelection(true)
      return
    }
    if (mod && e.key.toLowerCase() === 'v') {
      // let paste event handle
      return
    }
    if (e.key === 'Tab') {
      e.preventDefault()
      if (selection.focus.cell) {
        const { position, needsNewRow } = moveTableCellTab(doc, selection.focus, e.shiftKey ? -1 : 1)
        if (needsNewRow) {
          const before = cloneRange(selection)
          const afterRow = selection.focus.cell.row
          const afterPos: DocPosition = {
            sectionIndex: selection.focus.sectionIndex,
            blockIndex: selection.focus.blockIndex,
            offset: 0,
            cell: { row: afterRow + 1, cell: 0, para: 0 },
          }
          ;({ doc, history } = dispatch(
            doc,
            history,
            {
              type: 'insertRow',
              sectionIndex: selection.focus.sectionIndex,
              blockIndex: selection.focus.blockIndex,
              afterRow,
            },
            { before, after: { anchor: afterPos, focus: afterPos } },
          ))
          selection = { anchor: afterPos, focus: afterPos }
          preferredCol = 0
          selectAllLevel = 0
          render()
          return
        }
        setCaret(position, false)
        return
      }
      insertTab()
      return
    }
    if (e.key === 'Backspace') {
      e.preventDefault()
      deleteBackward()
      return
    }
    if (e.key === 'Delete') {
      e.preventDefault()
      deleteForward()
      return
    }
    if (e.key === 'Enter') {
      e.preventDefault()
      commitSplit()
      return
    }
    if (e.key === 'Home') {
      e.preventDefault()
      setCaret(moveHome(doc, selection.focus, mod), e.shiftKey)
      return
    }
    if (e.key === 'End') {
      e.preventDefault()
      setCaret(moveEnd(doc, selection.focus, mod), e.shiftKey)
      return
    }
    if (e.key === 'ArrowLeft') {
      e.preventDefault()
      setCaret(moveLeft(doc, selection.focus, mod), e.shiftKey)
      return
    }
    if (e.key === 'ArrowRight') {
      e.preventDefault()
      setCaret(moveRight(doc, selection.focus, mod), e.shiftKey)
      return
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault()
      moveCaretVertical(-1, e.shiftKey)
      return
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      moveCaretVertical(1, e.shiftKey)
      return
    }
  }

  const selectAll = () => {
    clearPendingMarks()
    const focus = selection.focus
    const block = focus.cell
      ? doc.sections[focus.sectionIndex]?.blocks[focus.blockIndex]
      : undefined

    // Word escalation: cell → table → document
    if (focus.cell && block?.type === 'table') {
      if (selectAllLevel < 1) {
        const tc = block.rows[focus.cell.row]?.cells[focus.cell.cell]
        if (tc && tc.blocks.length > 0) {
          const lastPara = tc.blocks.length - 1
          const startPos: DocPosition = {
            sectionIndex: focus.sectionIndex,
            blockIndex: focus.blockIndex,
            offset: 0,
            cell: { row: focus.cell.row, cell: focus.cell.cell, para: 0 },
          }
          const endPos: DocPosition = {
            sectionIndex: focus.sectionIndex,
            blockIndex: focus.blockIndex,
            offset: 0,
            cell: { row: focus.cell.row, cell: focus.cell.cell, para: lastPara },
          }
          endPos.offset = paragraphLength(getParagraph(doc, endPos))
          selection = { anchor: startPos, focus: endPos }
          selectAllLevel = 1
          preferredCol = endPos.offset
          caretBlink = true
          renderSelection()
          return
        }
      }
      if (selectAllLevel < 2) {
        let last: { row: number; cell: number; para: number } = { row: 0, cell: 0, para: 0 }
        for (let r = 0; r < block.rows.length; r++) {
          for (let c = 0; c < block.rows[r]!.cells.length; c++) {
            const n = block.rows[r]!.cells[c]!.blocks.length
            if (n > 0) last = { row: r, cell: c, para: n - 1 }
          }
        }
        const startPos: DocPosition = {
          sectionIndex: focus.sectionIndex,
          blockIndex: focus.blockIndex,
          offset: 0,
          cell: { row: 0, cell: 0, para: 0 },
        }
        const endPos: DocPosition = {
          sectionIndex: focus.sectionIndex,
          blockIndex: focus.blockIndex,
          offset: 0,
          cell: last,
        }
        endPos.offset = paragraphLength(getParagraph(doc, endPos))
        selection = { anchor: startPos, focus: endPos }
        selectAllLevel = 2
        preferredCol = endPos.offset
        caretBlink = true
        renderSelection()
        return
      }
    }

    const start = moveHome(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true)
    const end = moveEnd(doc, start, true)
    selection = { anchor: start, focus: end }
    selectAllLevel = 3
    renderSelection()
  }

  const clientToLocal = (e: PointerEvent) => {
    // Sticky canvas paints the viewport; map into document space via scroll.
    const rect = canvas.getBoundingClientRect()
    return {
      x: e.clientX - rect.left + host.scrollLeft,
      y: e.clientY - rect.top + host.scrollTop,
    }
  }

  const onPointerDown = (e: PointerEvent) => {
    host.focus()
    input.focus()
    clearPendingMarks()
    rectSelecting = false
    const { x, y } = clientToLocal(e)
    const hit = hitTestPoint(layout, x, y, zoom, PAGE_GAP_DEFAULT, pageColumns, host.clientWidth)
    if (!hit) return

    const now = performance.now()
    const sameSpot =
      lastClickPos &&
      lastClickPos.sectionIndex === hit.position.sectionIndex &&
      lastClickPos.blockIndex === hit.position.blockIndex &&
      Math.abs(lastClickPos.offset - hit.position.offset) <= 2 &&
      ((!lastClickPos.cell && !hit.position.cell) ||
        (lastClickPos.cell &&
          hit.position.cell &&
          lastClickPos.cell.row === hit.position.cell.row &&
          lastClickPos.cell.cell === hit.position.cell.cell &&
          lastClickPos.cell.para === hit.position.cell.para))
    if (now - lastClickAt < 400 && sameSpot) {
      clickCount += 1
    } else {
      clickCount = 1
    }
    lastClickAt = now
    lastClickPos = hit.position

    if (clickCount === 2 && e.button === 0) {
      // Double-click: select word
      const para = getParagraph(doc, hit.position)
      const text = paragraphPlainText(para)
      const { start, end } = wordBounds(text, hit.position.offset)
      selection = {
        anchor: { ...hit.position, offset: start },
        focus: { ...hit.position, offset: end },
      }
      preferredCol = end
      dragging = false
      renderSelection()
      return
    }
    if (clickCount >= 3 && e.button === 0) {
      // Triple-click: select paragraph
      const para = getParagraph(doc, hit.position)
      selection = {
        anchor: { ...hit.position, offset: 0 },
        focus: { ...hit.position, offset: paragraphLength(para) },
      }
      preferredCol = paragraphLength(para)
      dragging = false
      clickCount = 0
      renderSelection()
      return
    }

    // Click collapses selection to caret (Word/Docs). Drag after mousedown still extends.
    // Drag-move of selected text requires a deliberate drag (handled on pointermove threshold).
    if (!isCollapsed(selection) && e.button === 0 && clickCount === 1) {
      dragging = true
      dragAnchor = hit.position
      selection = { anchor: hit.position, focus: hit.position }
      preferredCol = hit.position.offset
      canvas.setPointerCapture(e.pointerId)
      renderSelection()
      return
    }

    dragging = true
    dragAnchor = hit.position
    selection = { anchor: hit.position, focus: hit.position }
    preferredCol = hit.position.offset
    canvas.setPointerCapture(e.pointerId)
    renderSelection()
  }

  const onPointerMove = (e: PointerEvent) => {
    const { x, y } = clientToLocal(e)
    const hit = hitTestPoint(layout, x, y, zoom, PAGE_GAP_DEFAULT, pageColumns, host.clientWidth)
    if (!hit) return
    if (dragMoving && dragSource) {
      selection = { ...dragSource, focus: hit.position }
      // Visual: show drop caret as focus only
      selection = { anchor: hit.position, focus: hit.position }
      paintOnly()
      return
    }
    if (!dragging) return
    let focus = hit.position
    // Drag-selecting inside a cell stays in that cell (Word character selection)
    const a = dragAnchor ?? selection.anchor
    if (a.cell) {
      const sameCell =
        focus.cell &&
        focus.sectionIndex === a.sectionIndex &&
        focus.blockIndex === a.blockIndex &&
        focus.cell.row === a.cell.row &&
        focus.cell.cell === a.cell.cell
      const sameTable =
        focus.cell &&
        focus.sectionIndex === a.sectionIndex &&
        focus.blockIndex === a.blockIndex
      if (sameCell) {
        // Still inside the anchor cell → ordinary character selection.
        rectSelecting = false
        selection = { anchor: a, focus }
        paintOnly()
        return
      }
      if (sameTable && focus.cell) {
        // Different cell in the same table → Word-style rectangular cell selection.
        // Anchor at the start of the top-left cell, focus at the end of the
        // bottom-right cell's last paragraph.
        rectSelecting = true
        const r0 = Math.min(a.cell.row, focus.cell.row)
        const r1 = Math.max(a.cell.row, focus.cell.row)
        const c0 = Math.min(a.cell.cell, focus.cell.cell)
        const c1 = Math.max(a.cell.cell, focus.cell.cell)
        const topLeft: DocPosition = {
          sectionIndex: a.sectionIndex,
          blockIndex: a.blockIndex,
          offset: 0,
          cell: { row: r0, cell: c0, para: 0 },
        }
        const bottomRight: DocPosition = {
          sectionIndex: a.sectionIndex,
          blockIndex: a.blockIndex,
          offset: 0,
          cell: { row: r1, cell: c1, para: lastCellPara(a.sectionIndex, a.blockIndex, r1, c1) },
        }
        bottomRight.offset = paragraphLength(getParagraph(doc, bottomRight))
        selection = { anchor: topLeft, focus: bottomRight }
        paintOnly()
        return
      }
      // Dragged out of the table entirely → clamp to the anchor cell edge.
      rectSelecting = false
      const edge: DocPosition = {
        sectionIndex: a.sectionIndex,
        blockIndex: a.blockIndex,
        offset: 0,
        cell: { ...a.cell },
      }
      const para = getParagraph(doc, edge)
      const len = paragraphLength(para)
      // Pick near/far edge of the cell paragraph from pointer vs anchor
      const after =
        focus.sectionIndex > a.sectionIndex ||
        (focus.sectionIndex === a.sectionIndex && focus.blockIndex > a.blockIndex)
      focus = { ...edge, offset: after ? len : 0 }
      selection = { anchor: a, focus }
      paintOnly()
      return
    }
    selection = { ...selection, focus }
    paintOnly()
  }

  const onPointerUp = (e: PointerEvent): void => {
    if (dragMoving && dragSource) {
      const { x, y } = clientToLocal(e)
      const hit = hitTestPoint(layout, x, y, zoom, PAGE_GAP_DEFAULT, pageColumns, host.clientWidth)
      if (hit) {
        const { start, end } = normalizeRange(dragSource)
        const text = extractPlainRange(doc, { anchor: start, focus: end })
        // Delete source then insert at drop if drop outside
        const drop = hit.position
        const droppingInside =
          (drop.sectionIndex > start.sectionIndex ||
            (drop.sectionIndex === start.sectionIndex && drop.blockIndex > start.blockIndex) ||
            (drop.sectionIndex === start.sectionIndex &&
              drop.blockIndex === start.blockIndex &&
              drop.offset >= start.offset)) &&
          (drop.sectionIndex < end.sectionIndex ||
            (drop.sectionIndex === end.sectionIndex && drop.blockIndex < end.blockIndex) ||
            (drop.sectionIndex === end.sectionIndex &&
              drop.blockIndex === end.blockIndex &&
              drop.offset <= end.offset))
        if (!droppingInside && text) {
          const before = cloneRange(dragSource)
          let insertAt = drop
          if (
            drop.sectionIndex === start.sectionIndex &&
            drop.blockIndex === start.blockIndex &&
            drop.offset > end.offset
          ) {
            insertAt = { ...drop, offset: drop.offset - (end.offset - start.offset) }
          } else if (
            drop.sectionIndex === start.sectionIndex &&
            drop.blockIndex === start.blockIndex &&
            drop.offset > start.offset
          ) {
            insertAt = start
          }
          const after: DocRange = {
            anchor: insertAt,
            focus: { ...insertAt, offset: insertAt.offset + text.length },
          }
          ;({ doc, history } = dispatchBatch(
            doc,
            history,
            [
              { type: 'deleteRange', range: { anchor: start, focus: end } },
              { type: 'insertText', position: insertAt, text },
            ],
            { before, after },
          ))
          selection = after
          render()
        } else {
          selection = dragSource
          renderSelection()
        }
      }
      dragMoving = false
      dragSource = undefined
      dragAnchor = undefined
      return
    }
    dragging = false
    dragAnchor = undefined
  }

  let editorActive = true
  const onDocKeyDown = (e: KeyboardEvent) => {
    if (!editorActive) return
    // Don't steal keys from real form fields outside the editor
    const target = e.target as HTMLElement | null
    if (
      target &&
      target !== input.el &&
      target !== host &&
      (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable)
    ) {
      return
    }
    // Textarea already handles keys — skip so Backspace/Delete aren't applied 2–3×
    if (target === input.el) return
    // Toolbar / body: route shortcuts into the editor once
    if (target === document.body || target?.closest?.('.toolbar')) {
      onKeyDown(e)
      if (e.defaultPrevented) input.focus()
    }
  }

  // Focus lives on the textarea — single keydown path avoids triple-delete
  input.el.addEventListener('keydown', onKeyDown)
  document.addEventListener('keydown', onDocKeyDown)
  host.addEventListener('focusin', () => {
    editorActive = true
  })
  const onPasteListener = (e: ClipboardEvent) => {
    void onPaste(e)
  }
  host.addEventListener('paste', onPasteListener)
  host.addEventListener('copy', onCopy)
  host.addEventListener('cut', onCut)
  canvas.addEventListener('pointerdown', onPointerDown)
  canvas.addEventListener('pointermove', onPointerMove)
  canvas.addEventListener('pointerup', onPointerUp)
  host.addEventListener('focus', () => input.focus())
  const onScroll = () => paintOnly()
  host.addEventListener('scroll', onScroll, { passive: true })
  const onResize = () => {
    if (zoomMode !== 'percent') render()
    else paintOnly()
  }
  window.addEventListener('resize', onResize)

  const blink = window.setInterval(() => {
    caretBlink = !caretBlink
    paintOnly()
  }, 530)

  loadImages()
  render()

  return {
    getDocument: () => doc,
    setDocument: (next) => {
      doc = next
      history = createHistory()
      clearPendingMarks()
      selection = {
        anchor: moveHome(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true),
        focus: moveHome(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true),
      }
      loadImages()
      render()
    },
    loadBytes: (bytes, format) => {
      doc = loadDocument(bytes, format)
      history = createHistory()
      clearPendingMarks()
      const home = moveHome(doc, { sectionIndex: 0, blockIndex: 0, offset: 0 }, true)
      selection = { anchor: home, focus: home }
      loadImages()
      render()
    },
    saveBytes: (format) => saveDocument(doc, format),
    focus: () => {
      host.focus()
      input.focus()
    },
    destroy: () => {
      window.clearInterval(blink)
      if (fullLayoutTimer) clearTimeout(fullLayoutTimer)
      if (notifyTimer) clearTimeout(notifyTimer)
      window.removeEventListener('resize', onResize)
      host.removeEventListener('scroll', onScroll)
      for (const url of objectUrls) URL.revokeObjectURL(url)
      input.el.removeEventListener('keydown', onKeyDown)
      document.removeEventListener('keydown', onDocKeyDown)
      host.removeEventListener('paste', onPasteListener)
      host.removeEventListener('copy', onCopy)
      host.removeEventListener('cut', onCut)
      canvas.removeEventListener('pointerdown', onPointerDown)
      canvas.removeEventListener('pointermove', onPointerMove)
      canvas.removeEventListener('pointerup', onPointerUp)
      input.destroy()
      a11y.destroy()
      spacer.remove()
    },
    getSelection: () => selection,
    undo: () => {
      if (!canUndo(history)) return
      applyHistory('undo')
    },
    redo: () => {
      if (!canRedo(history)) return
      applyHistory('redo')
    },
    setZoom: (z) => {
      zoomMode = 'percent'
      zoom = Math.min(4, Math.max(0.25, z))
      input.focus()
      paintOnly()
    },
    getZoom: () => zoom,
    setZoomMode: (mode) => {
      zoomMode = mode
      input.focus()
      paintOnly()
    },
    getZoomMode: () => zoomMode,
    bold: () => applyMark({ bold: true }),
    italic: () => applyMark({ italic: true }),
    underline: () => applyMark({ underline: true }),
    strike: () => applyMark({ strike: true }),
    align: applyAlign,
    bulletList: () => applyList({ numId: '1', ilvl: 0 }),
    numberedList: () => applyList({ numId: '2', ilvl: 0 }),
    clearList: () => applyList(null),
    find: (opts) => findAll(doc, opts),
    findNext: (opts) => {
      const m = findNext(doc, opts, selection.focus)
      if (m) {
        clearPendingMarks()
        selection = m.range
        renderSelection()
      }
      return m
    },
    replace: (opts, replacement) => {
      const m = findNext(doc, opts, selection.focus) ?? findAll(doc, opts)[0]
      if (!m) return false
      const result = replaceMatch(doc, m, replacement)
      const { start, end } = normalizeRange(m.range)
      const after: DocRange = {
        anchor: start,
        focus: { ...start, offset: start.offset + replacement.length },
      }
      const ops: EditorOp[] = [{ type: 'deleteRange', range: { anchor: start, focus: end } }]
      if (replacement) {
        ops.push({ type: 'insertText', position: start, text: replacement })
      }
      void result
      commitOps(ops, after)
      return true
    },
    replaceAll: (opts, replacement) => {
      const matches = findAll(doc, opts)
      const before = cloneRange(selection)
      const ops: EditorOp[] = []
      for (let i = matches.length - 1; i >= 0; i--) {
        const m = matches[i]!
        const { start, end } = normalizeRange(m.range)
        ops.push({ type: 'deleteRange', range: { anchor: start, focus: end } })
        if (replacement) {
          ops.push({ type: 'insertText', position: start, text: replacement })
        }
      }
      if (ops.length) {
        ;({ doc, history } = dispatchBatch(doc, history, ops, { before, after: before }))
        render()
      }
      return matches.length
    },
    selectAll,
  }
}

