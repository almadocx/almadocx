import type {
  Document,
  HeaderFooter,
  Paragraph,
  Table,
} from '../model/types.js'
import { resolveListMarkers, type ResolvedListMarker } from '../model/numbering.js'
import { resolveParagraphProps, resolveRunProps } from '../styles/cascade.js'
import { twipsToPx, pxToTwips } from '../util/units.js'
import { createApproximateMeasurer, fontCss, type TextMeasurer } from './measure.js'
import { inlineToChar } from '../model/position.js'
import { resolveFontFamily } from '../fonts/substitutes.js'
import { nextTabStopTwips } from './tabs.js'

export interface LayoutGlyphRun {
  text: string
  x: number
  width: number
  font: string
  color: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  startOffset: number
  image?: { mediaId: string; width: number; height: number; alt?: string }
  tabLeader?: 'none' | 'dot' | 'dash' | 'underscore'
}

export interface LayoutLine {
  y: number
  height: number
  baseline: number
  runs: LayoutGlyphRun[]
  startOffset: number
  endOffset: number
}

export interface LayoutParagraph {
  kind: 'paragraph'
  blockIndex: number
  sectionIndex: number
  x: number
  y: number
  width: number
  height: number
  lines: LayoutLine[]
  marker?: { text: string; x: number; y: number; font: string; color: string }
  /** When laid out inside a table cell */
  cell?: { row: number; cell: number; para: number }
}

export interface LayoutTableCell {
  x: number
  y: number
  width: number
  height: number
  paragraphs: LayoutParagraph[]
  shading?: string
  borderColor: string
  rowIndex: number
  cellIndex: number
}

export interface LayoutTable {
  kind: 'table'
  blockIndex: number
  sectionIndex: number
  x: number
  y: number
  width: number
  height: number
  cells: LayoutTableCell[]
  startRow: number
  endRow: number
}

export interface LayoutHeaderFooter {
  paragraphs: LayoutParagraph[]
  pageNumberText?: string
}

export type LayoutBlock = LayoutParagraph | LayoutTable

export interface LayoutPage {
  index: number
  width: number
  height: number
  blocks: LayoutBlock[]
  /** @deprecated use blocks — kept for Phase 1 callers */
  paragraphs: LayoutParagraph[]
  header?: LayoutHeaderFooter
  footer?: LayoutHeaderFooter
}

export interface LayoutResult {
  pages: LayoutPage[]
}

export interface LayoutOptions {
  measurer?: TextMeasurer
  dpi?: number
}

interface Ctx {
  measurer: TextMeasurer
  dpi: number
}

export function layoutDocument(doc: Document, options: LayoutOptions = {}): LayoutResult {
  const dpi = options.dpi ?? 96
  const measurer = options.measurer ?? createApproximateMeasurer()
  const ctx: Ctx = { measurer, dpi }
  const pages: LayoutPage[] = []

  for (let si = 0; si < doc.sections.length; si++) {
    const section = doc.sections[si]!
    const pageW = twipsToPx(section.properties.pageSize.width, dpi)
    const pageH = twipsToPx(section.properties.pageSize.height, dpi)
    const margin = section.properties.margins
    const contentX = twipsToPx(margin.left, dpi)
    const headerBand = twipsToPx(margin.header, dpi)
    const footerBand = twipsToPx(margin.footer, dpi)
    const contentTop = Math.max(twipsToPx(margin.top, dpi), headerBand + 8)
    const contentBottom = pageH - Math.max(twipsToPx(margin.bottom, dpi), footerBand + 8)
    const contentWidth = pageW - twipsToPx(margin.left, dpi) - twipsToPx(margin.right, dpi)
    const markers = resolveListMarkers(doc, si)

    let pageIndex = 0
    let cursorY = contentTop
    let currentBlocks: LayoutBlock[] = []

    const makePage = (blocks: LayoutBlock[], index: number): LayoutPage => {
      const paragraphs = blocks.filter((b): b is LayoutParagraph => b.kind === 'paragraph')
      const page: LayoutPage = {
        index,
        width: pageW,
        height: pageH,
        blocks,
        paragraphs,
      }
      const useFirst = section.properties.titlePage && index === 0
      const headerSrc = useFirst ? section.headerFirst ?? section.header : section.header
      const footerSrc = useFirst ? section.footerFirst ?? section.footer : section.footer
      if (headerSrc) {
        page.header = layoutHeaderFooter(
          doc,
          headerSrc,
          contentX,
          headerBand * 0.35,
          contentWidth,
          ctx,
          index + 1,
        )
      } else {
        // Do not invent a header page number — only paint what the document defines
        page.header = { paragraphs: [] }
      }
      if (footerSrc) {
        page.footer = layoutHeaderFooter(
          doc,
          footerSrc,
          contentX,
          pageH - footerBand * 0.85,
          contentWidth,
          ctx,
          index + 1,
        )
      } else {
        page.footer = { paragraphs: [] }
      }
      return page
    }

    const flushPage = () => {
      pages.push(makePage(currentBlocks, pageIndex))
      pageIndex += 1
      currentBlocks = []
      cursorY = contentTop
    }

    for (let bi = 0; bi < section.blocks.length; bi++) {
      const block = section.blocks[bi]!

      if (block.type === 'paragraph') {
        const pProps = resolveParagraphProps(doc, block)
        if (pProps.pageBreakBefore && currentBlocks.length > 0) flushPage()

        const marker = markers.get(bi)
        const laid = layoutParagraphBlock(
          doc,
          block,
          si,
          bi,
          contentX,
          cursorY,
          contentWidth,
          ctx,
          marker,
        )

        // Widow/orphan: avoid leaving a single line alone when keepLines or widowControl
        const widowControl = pProps.widowControl !== false
        const keepLines = pProps.keepLines === true
        const keepNext = pProps.keepNext === true

        if (cursorY > contentTop && laid.y + laid.height > contentBottom) {
          // If keepNext with following block, try to keep together by flushing earlier
          if (keepNext || keepLines) {
            flushPage()
            reflowParagraphY(laid, contentTop)
          } else if (widowControl && laid.lines.length >= 3) {
            // Split with at least 2 lines on each side when possible
            const chunks = splitParagraphWidowAware(laid, contentTop, contentBottom)
            for (let c = 0; c < chunks.length; c++) {
              const chunk = chunks[c]!
              // fit===0 reflows the whole para to contentTop — must start a new page
              // or it will paint over content already on this page.
              const startsNewPage =
                c > 0 || (currentBlocks.length > 0 && chunk.y <= contentTop + 1)
              if (startsNewPage && currentBlocks.length > 0) {
                flushPage()
                reflowParagraphY(chunk, contentTop)
              }
              currentBlocks.push(chunk)
              cursorY = chunk.y + chunk.height
            }
            continue
          } else {
            flushPage()
            reflowParagraphY(laid, contentTop)
          }
        }

        if (laid.height > contentBottom - contentTop) {
          const chunks = splitParagraphAcrossPages(laid, contentTop, contentBottom)
          for (let c = 0; c < chunks.length; c++) {
            const chunk = chunks[c]!
            if (c > 0 || (currentBlocks.length > 0 && chunk.y <= contentTop + 1)) {
              if (currentBlocks.length > 0) flushPage()
              reflowParagraphY(chunk, contentTop)
            }
            currentBlocks.push(chunk)
            cursorY = chunk.y + chunk.height
          }
        } else {
          currentBlocks.push(laid)
          cursorY = laid.y + laid.height
        }
      } else if (block.type === 'table') {
        let rowStart = 0
        while (rowStart < block.rows.length) {
          if (cursorY > contentBottom - 24 && currentBlocks.length > 0) flushPage()

          // Greedily take as many rows as fit on this page
          let lo = rowStart
          let hi = block.rows.length
          let best = rowStart
          while (lo < hi) {
            const mid = Math.floor((lo + hi) / 2)
            const probe = layoutTableRows(
              doc,
              block,
              si,
              bi,
              contentX,
              cursorY,
              contentWidth,
              ctx,
              rowStart,
              mid + 1,
            )
            if (cursorY + probe.height <= contentBottom || mid === rowStart) {
              best = mid
              lo = mid + 1
            } else {
              hi = mid
            }
          }
          const endExclusive = Math.max(best + 1, rowStart + 1)
          let laid = layoutTableRows(
            doc,
            block,
            si,
            bi,
            contentX,
            cursorY,
            contentWidth,
            ctx,
            rowStart,
            endExclusive,
          )
          if (cursorY > contentTop && cursorY + laid.height > contentBottom && currentBlocks.length > 0) {
            flushPage()
            laid = layoutTableRows(
              doc,
              block,
              si,
              bi,
              contentX,
              contentTop,
              contentWidth,
              ctx,
              rowStart,
              endExclusive,
            )
          }
          currentBlocks.push(laid)
          cursorY = laid.y + laid.height
          rowStart = endExclusive
        }
      }
    }

    pages.push(makePage(currentBlocks, pageIndex))
  }

  return { pages }
}

function layoutHeaderFooter(
  doc: Document,
  hf: HeaderFooter,
  x: number,
  y: number,
  width: number,
  ctx: Ctx,
  pageNumber: number,
): LayoutHeaderFooter {
  const paragraphs: LayoutParagraph[] = []
  let cursor = y
  let pageNumberText: string | undefined
  for (let i = 0; i < hf.blocks.length; i++) {
    const p = hf.blocks[i]!
    // Replace PAGE field-like text "{{PAGE}}" if present
    const runs = p.runs.map((r) => {
      if (r.content.type === 'text' && r.content.text.includes('{{PAGE}}')) {
        return {
          ...r,
          content: { type: 'text' as const, text: r.content.text.replaceAll('{{PAGE}}', String(pageNumber)) },
        }
      }
      return r
    })
    const para = { ...p, runs }
    const laid = layoutParagraphBlock(doc, para, 0, i, x, cursor, width, ctx)
    paragraphs.push(laid)
    cursor = laid.y + laid.height
    for (const run of p.runs) {
      if (run.content.type === 'text' && run.content.text.includes('{{PAGE}}')) {
        pageNumberText = String(pageNumber)
      }
    }
  }
  if (!pageNumberText && paragraphs.length === 0) pageNumberText = String(pageNumber)
  return { paragraphs, ...(pageNumberText !== undefined ? { pageNumberText } : {}) }
}

function reflowParagraphY(laid: LayoutParagraph, y: number): void {
  const dy = y - laid.y
  laid.y = y
  for (const line of laid.lines) line.y += dy
  if (laid.marker) laid.marker.y += dy
}

function splitParagraphAcrossPages(
  para: LayoutParagraph,
  contentTop: number,
  contentBottom: number,
): LayoutParagraph[] {
  const out: LayoutParagraph[] = []
  let current: LayoutParagraph = { ...para, y: contentTop, lines: [], height: 0 }
  let y = contentTop
  for (const line of para.lines) {
    if (y + line.height > contentBottom && current.lines.length > 0) {
      current.height = y - current.y
      out.push(current)
      y = contentTop
      current = { ...para, y, lines: [], height: 0 }
    }
    current.lines.push({ ...line, y })
    y += line.height
  }
  current.height = y - current.y
  out.push(current)
  return out
}

function splitParagraphWidowAware(
  para: LayoutParagraph,
  contentTop: number,
  contentBottom: number,
): LayoutParagraph[] {
  const avail = contentBottom - (para.y > contentTop ? para.y : contentTop)
  let fit = 0
  let used = 0
  for (const line of para.lines) {
    if (used + line.height > avail) break
    used += line.height
    fit += 1
  }
  // Keep at least 2 lines on the next page when splitting (orphan control)
  if (fit > 0 && para.lines.length - fit < 2 && para.lines.length >= 4) {
    fit = Math.max(2, para.lines.length - 2)
  }
  // Keep at least 2 lines on current page (widow control)
  if (fit === 1 && para.lines.length >= 3) {
    fit = 0 // push all to next page
  }
  if (fit === 0) {
    const moved = { ...para }
    reflowParagraphY(moved, contentTop)
    return [moved]
  }
  if (fit >= para.lines.length) return [para]

  const firstLines = para.lines.slice(0, fit)
  const secondLines = para.lines.slice(fit)
  const first: LayoutParagraph = {
    ...para,
    lines: firstLines,
    height: firstLines.reduce((h, l) => h + l.height, 0),
  }
  const second: LayoutParagraph = {
    ...para,
    y: contentTop,
    lines: [],
    height: 0
  }
  let y = contentTop
  for (const line of secondLines) {
    second.lines.push({ ...line, y })
    y += line.height
  }
  second.height = y - contentTop
  return [first, second]
}

function layoutParagraphBlock(
  doc: Document,
  paragraph: Paragraph,
  sectionIndex: number,
  blockIndex: number,
  x: number,
  y: number,
  maxWidth: number,
  ctx: Ctx,
  marker?: ResolvedListMarker,
  cellPath?: { row: number; cell: number; para: number },
): LayoutParagraph {
  const pProps = resolveParagraphProps(doc, paragraph)
  let indentLeft = twipsToPx(pProps.indentLeft ?? 0, ctx.dpi)
  const indentRight = twipsToPx(pProps.indentRight ?? 0, ctx.dpi)
  let indentFirst = twipsToPx(pProps.indentFirstLine ?? 0, ctx.dpi)
  const spacingBefore = twipsToPx(pProps.spacingBefore ?? 0, ctx.dpi)
  const spacingAfter = twipsToPx(pProps.spacingAfter ?? 0, ctx.dpi)
  const lineMult =
    pProps.lineSpacingRule === 'auto' || pProps.lineSpacingRule === undefined
      ? (pProps.lineSpacing ?? 1.15)
      : 1.15

  let markerLayout: LayoutParagraph['marker']
  if (marker) {
    indentLeft = Math.max(indentLeft, twipsToPx(marker.indentLeft, ctx.dpi))
    indentFirst = -twipsToPx(marker.hanging, ctx.dpi)
    const fontFamily = resolveFontFamily(marker.fontFamily ?? doc.styles.docDefaults.character.fontFamily)
    const fontSize = doc.styles.docDefaults.character.fontSizePt ?? 11
    const font = fontCss(fontFamily, fontSize, false, false, ctx.dpi)
    const metrics = ctx.measurer.measure(marker.text, font)
    markerLayout = {
      text: marker.text,
      x: x + indentLeft - twipsToPx(marker.hanging, ctx.dpi),
      y: y + spacingBefore,
      font,
      color: doc.styles.docDefaults.character.color ?? '#000000',
    }
    void metrics
  }

  const width = Math.max(0, maxWidth - indentLeft - indentRight)
  const lines: LayoutLine[] = []
  let lineRuns: LayoutGlyphRun[] = []
  let lineWidth = 0
  let lineAscent = 0
  let lineDescent = 0
  let lineStartOffset = 0
  let offset = 0
  let isFirstLine = true

  const flushLine = (endOffset: number) => {
    const height = Math.max(1, (lineAscent + lineDescent) * lineMult)
    const baseline = lineAscent
    const firstExtra = isFirstLine ? indentFirst : 0
    const available = width - Math.max(0, firstExtra)
    let originX = x + indentLeft + (isFirstLine ? Math.max(0, indentFirst) : 0)
    if (pProps.alignment === 'center') {
      originX += Math.max(0, (available - lineWidth) / 2)
    } else if (pProps.alignment === 'right') {
      originX += Math.max(0, available - lineWidth)
    }
    const runs = lineRuns.map((r) => ({ ...r, x: r.x + originX }))
    lines.push({
      y: 0,
      height,
      baseline,
      runs,
      startOffset: lineStartOffset,
      endOffset,
    })
    lineRuns = []
    lineWidth = 0
    lineAscent = 0
    lineDescent = 0
    lineStartOffset = endOffset
    isFirstLine = false
  }

  for (const run of paragraph.runs) {
    const resolved = resolveRunProps(doc, paragraph, run)
    const font = fontCss(
      resolveFontFamily(resolved.fontFamily),
      resolved.fontSizePt ?? 11,
      resolved.bold,
      resolved.italic,
      ctx.dpi,
    )
    const color = resolved.color ?? '#000000'

    if (run.content.type === 'image') {
      const w = twipsToPx(run.content.widthTwips, ctx.dpi)
      const h = twipsToPx(run.content.heightTwips, ctx.dpi)
      const avail = width - (isFirstLine ? Math.max(0, indentFirst) : 0)
      if (lineWidth + w > avail && lineRuns.length > 0) flushLine(offset)
      const imageRun: LayoutGlyphRun = {
        text: '',
        x: lineWidth,
        width: w,
        font,
        color,
        startOffset: offset,
        image: {
          mediaId: run.content.mediaId,
          width: w,
          height: h,
          ...(run.content.alt !== undefined ? { alt: run.content.alt } : {}),
        },
      }
      lineRuns.push(imageRun)
      lineWidth += w
      lineAscent = Math.max(lineAscent, h * 0.85)
      lineDescent = Math.max(lineDescent, h * 0.15)
      offset += 1
      continue
    }

    const text = inlineToChar(run)

    if (run.content.type === 'tab') {
      const contentWidthTwips = pxToTwips(width, ctx.dpi)
      const currentTwips = pxToTwips(lineWidth, ctx.dpi)
      const stop = nextTabStopTwips(currentTwips, pProps.tabs, contentWidthTwips)
      const targetX = twipsToPx(stop.position, ctx.dpi)
      const w = Math.max(twipsToPx(36, ctx.dpi), targetX - lineWidth)
      const metrics = ctx.measurer.measure(' ', font)
      lineRuns.push({
        text: '\t',
        x: lineWidth,
        width: w,
        font,
        color,
        startOffset: offset,
        tabLeader: stop.leader ?? 'none',
        ...(resolved.bold ? { bold: true } : {}),
        ...(resolved.italic ? { italic: true } : {}),
      })
      lineWidth += w
      lineAscent = Math.max(lineAscent, metrics.ascent)
      lineDescent = Math.max(lineDescent, metrics.descent)
      offset += 1
      continue
    }

    if (run.content.type !== 'text') {
      const metrics = ctx.measurer.measure(' ', font)
      const w = metrics.width
      const avail = width - (isFirstLine ? Math.max(0, indentFirst) : 0)
      if (lineWidth + w > avail && lineRuns.length > 0) flushLine(offset)
      lineRuns.push({
        text,
        x: lineWidth,
        width: w,
        font,
        color,
        startOffset: offset,
        ...(resolved.bold ? { bold: true } : {}),
        ...(resolved.italic ? { italic: true } : {}),
        ...(resolved.underline ? { underline: true } : {}),
        ...(resolved.strike ? { strike: true } : {}),
      })
      lineWidth += w
      lineAscent = Math.max(lineAscent, metrics.ascent)
      lineDescent = Math.max(lineDescent, metrics.descent)
      offset += text.length
      if (run.content.type === 'break' && run.content.breakType === 'page') flushLine(offset)
      continue
    }

    const words = splitKeepSep(text)
    for (const word of words) {
      const metrics = ctx.measurer.measure(word, font)
      const avail = Math.max(1, width - (isFirstLine ? Math.max(0, indentFirst) : 0))
      if (lineWidth + metrics.width > avail && lineRuns.length > 0 && word.trim() !== '') {
        flushLine(offset)
      }
      // Hard-break unbreakable runs that still exceed the line
      let remaining = word
      while (remaining.length > 0) {
        const remMetrics = ctx.measurer.measure(remaining, font)
        const lineAvail = Math.max(1, width - (isFirstLine ? Math.max(0, indentFirst) : 0))
        if (lineWidth + remMetrics.width <= lineAvail || remaining.length === 1) {
          lineRuns.push({
            text: remaining,
            x: lineWidth,
            width: remMetrics.width,
            font,
            color,
            startOffset: offset,
            ...(resolved.bold ? { bold: true } : {}),
            ...(resolved.italic ? { italic: true } : {}),
            ...(resolved.underline ? { underline: true } : {}),
            ...(resolved.strike ? { strike: true } : {}),
          })
          lineWidth += remMetrics.width
          lineAscent = Math.max(lineAscent, remMetrics.ascent)
          lineDescent = Math.max(lineDescent, remMetrics.descent)
          offset += remaining.length
          remaining = ''
          break
        }
        // Binary search how many chars fit
        let lo = 1
        let hi = remaining.length
        let fit = 1
        while (lo <= hi) {
          const mid = Math.floor((lo + hi) / 2)
          const slice = remaining.slice(0, mid)
          const w = ctx.measurer.measure(slice, font).width
          if (lineWidth + w <= lineAvail) {
            fit = mid
            lo = mid + 1
          } else {
            hi = mid - 1
          }
        }
        if (fit < 1) fit = 1
        const piece = remaining.slice(0, fit)
        const pieceMetrics = ctx.measurer.measure(piece, font)
        lineRuns.push({
          text: piece,
          x: lineWidth,
          width: pieceMetrics.width,
          font,
          color,
          startOffset: offset,
          ...(resolved.bold ? { bold: true } : {}),
          ...(resolved.italic ? { italic: true } : {}),
          ...(resolved.underline ? { underline: true } : {}),
          ...(resolved.strike ? { strike: true } : {}),
        })
        lineWidth += pieceMetrics.width
        lineAscent = Math.max(lineAscent, pieceMetrics.ascent)
        lineDescent = Math.max(lineDescent, pieceMetrics.descent)
        offset += piece.length
        remaining = remaining.slice(fit)
        if (remaining.length > 0) flushLine(offset)
      }
    }
  }

  if (lineRuns.length > 0 || lines.length === 0) flushLine(offset)

  let cursor = y + spacingBefore
  for (const line of lines) {
    line.y = cursor
    cursor += line.height
  }
  cursor += spacingAfter

  if (markerLayout) {
    markerLayout.y = y + spacingBefore
  }

  return {
    kind: 'paragraph',
    sectionIndex,
    blockIndex,
    x: x + indentLeft,
    y: y + spacingBefore,
    width,
    // Height is measured from the content top (first baseline band), not from the
    // pre-spacing origin — so laid.y + laid.height advances correctly without
    // double-counting spacingBefore.
    height: cursor - (y + spacingBefore),
    lines,
    ...(markerLayout ? { marker: markerLayout } : {}),
    ...(cellPath ? { cell: cellPath } : {}),
  }
}

function layoutTableRows(
  doc: Document,
  table: Table,
  sectionIndex: number,
  blockIndex: number,
  x: number,
  y: number,
  maxWidth: number,
  ctx: Ctx,
  startRow: number,
  endRowExclusive: number,
): LayoutTable {
  const sliced: Table = {
    ...table,
    rows: table.rows.slice(startRow, endRowExclusive),
  }
  const laid = layoutTableBlock(doc, sliced, sectionIndex, blockIndex, x, y, maxWidth, ctx, 0)
  for (const c of laid.cells) {
    c.rowIndex = startRow + c.rowIndex
    for (const p of c.paragraphs) {
      if (p.cell) p.cell = { ...p.cell, row: c.rowIndex }
    }
  }
  laid.startRow = startRow
  laid.endRow = endRowExclusive - 1
  return laid
}

function layoutTableBlock(
  doc: Document,
  table: Table,
  sectionIndex: number,
  blockIndex: number,
  x: number,
  y: number,
  maxWidth: number,
  ctx: Ctx,
  startRow = 0,
): LayoutTable {
  const colCount = Math.max(...table.rows.map((r) => r.cells.reduce((n, c) => n + (c.props.gridSpan ?? 1), 0)), 1)
  const rawGrid =
    table.gridCols && table.gridCols.length > 0
      ? table.gridCols.map((t) => twipsToPx(t, ctx.dpi))
      : []
  const grid =
    rawGrid.length >= colCount
      ? rawGrid.slice(0, colCount)
      : Array.from({ length: colCount }, (_, i) => rawGrid[i] ?? maxWidth / colCount)

  const totalGrid = grid.reduce((a, b) => a + b, 0) || 1
  const scale = maxWidth / totalGrid
  const colWidths = grid.map((w) => w * scale)

  const cells: LayoutTableCell[] = []
  let rowY = y
  const borderColor = table.props.borders?.insideH ?? table.props.borders?.top ?? '#c8c2b8'
  let lastRow = startRow - 1

  const DEFAULT_PAD = 4
  const padPx = (twips: number | undefined): number =>
    twips !== undefined ? twipsToPx(twips, ctx.dpi) : DEFAULT_PAD

  interface CellMeta {
    cell: LayoutTableCell
    col: number
    span: number
    vAlign?: 'top' | 'center' | 'bottom'
    marginTop: number
    marginBottom: number
    contentHeight: number
    isRestart: boolean
  }
  // Row-offset-indexed bookkeeping for vertical-merge expansion.
  const rowHeights: number[] = []
  const rowContinueCols: Set<number>[] = []
  const allMeta: CellMeta[] = []

  for (let rowIndex = startRow; rowIndex < table.rows.length; rowIndex++) {
    const row = table.rows[rowIndex]!
    let col = 0
    let rowHeight = 0
    const continueCols = new Set<number>()
    const rowMeta: CellMeta[] = []
    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
      const cell = row.cells[cellIndex]!
      const span = cell.props.gridSpan ?? 1
      if (cell.props.vMerge === 'continue') {
        continueCols.add(col)
        col += span
        continue
      }
      let cellX = x
      for (let i = 0; i < col; i++) cellX += colWidths[i] ?? 0
      let cellW = 0
      for (let i = col; i < col + span; i++) cellW += colWidths[i] ?? 0

      const marginTop = padPx(cell.props.margin?.top)
      const marginBottom = padPx(cell.props.margin?.bottom)
      const marginLeft = padPx(cell.props.margin?.left)
      const marginRight = padPx(cell.props.margin?.right)

      const paras: LayoutParagraph[] = []
      let cy = rowY + marginTop
      for (let pi = 0; pi < cell.blocks.length; pi++) {
        const p = cell.blocks[pi]!
        if (p.type !== 'paragraph') continue
        const laid = layoutParagraphBlock(
          doc,
          p,
          sectionIndex,
          blockIndex,
          cellX + marginLeft,
          cy,
          Math.max(8, cellW - marginLeft - marginRight),
          ctx,
          undefined,
          { row: rowIndex, cell: cellIndex, para: pi },
        )
        paras.push(laid)
        cy = laid.y + laid.height
      }
      const contentHeight = cy - (rowY + marginTop)
      const cellH = Math.max(
        contentHeight + marginTop + marginBottom,
        twipsToPx(row.props.heightTwips ?? 0, ctx.dpi),
        24,
      )
      rowHeight = Math.max(rowHeight, cellH)
      const laidCell: LayoutTableCell = {
        x: cellX,
        y: rowY,
        width: cellW,
        height: cellH,
        paragraphs: paras,
        borderColor,
        rowIndex,
        cellIndex,
      }
      if (cell.props.shading) laidCell.shading = cell.props.shading
      const meta: CellMeta = {
        cell: laidCell,
        col,
        span,
        marginTop,
        marginBottom,
        contentHeight,
        isRestart: cell.props.vMerge === 'restart',
        ...(cell.props.vAlign ? { vAlign: cell.props.vAlign } : {}),
      }
      rowMeta.push(meta)
      allMeta.push(meta)
      col += span
    }
    for (const m of rowMeta) m.cell.height = rowHeight
    cells.push(...rowMeta.map((m) => m.cell))
    rowHeights.push(rowHeight)
    rowContinueCols.push(continueCols)
    rowY += rowHeight
    lastRow = rowIndex
  }

  // Vertical merge: expand restart cells to cover subsequent `continue` rows in
  // the same starting column. Continue cells are not laid out, so their row
  // heights are added onto the originating restart cell's box.
  for (const m of allMeta) {
    if (!m.isRestart) continue
    const startOffset = m.cell.rowIndex - startRow
    for (let r = startOffset + 1; r < rowHeights.length; r++) {
      if (!rowContinueCols[r]!.has(m.col)) break
      m.cell.height += rowHeights[r]!
    }
  }

  // Vertical alignment: offset cell paragraphs within the (possibly merged) box.
  for (const m of allMeta) {
    if (!m.vAlign || m.vAlign === 'top') continue
    const avail = m.cell.height - m.marginTop - m.marginBottom
    const slack = avail - m.contentHeight
    if (slack <= 0.5) continue
    const dy = m.vAlign === 'center' ? slack / 2 : slack
    for (let i = 0; i < m.cell.paragraphs.length; i++) {
      m.cell.paragraphs[i] = shiftParagraphY(m.cell.paragraphs[i]!, m.cell.paragraphs[i]!.y + dy)
    }
  }

  return {
    kind: 'table',
    sectionIndex,
    blockIndex,
    x,
    y,
    width: maxWidth,
    height: rowY - y + 8,
    cells,
    startRow,
    endRow: lastRow,
  }
}

function splitKeepSep(text: string): string[] {
  const out: string[] = []
  const re = /(\s+|\S+)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(text))) out.push(m[1]!)
  return out.length ? out : ['']
}

export interface PatchLayoutResult {
  layout: LayoutResult
  /** True when height changed enough that pagination may be wrong — schedule a full layout. */
  needsFullLayout: boolean
}

/**
 * Fast path for typing: re-measure a single body paragraph in place.
 * Returns needsFullLayout when height changes or the block is split / in a table.
 */
export function patchLayoutParagraph(
  doc: Document,
  layout: LayoutResult,
  sectionIndex: number,
  blockIndex: number,
  options: LayoutOptions = {},
): PatchLayoutResult {
  const dpi = options.dpi ?? 96
  const measurer = options.measurer ?? createApproximateMeasurer()
  const ctx: Ctx = { measurer, dpi }

  const section = doc.sections[sectionIndex]
  const block = section?.blocks[blockIndex]
  if (!block || block.type !== 'paragraph') {
    return { layout, needsFullLayout: true }
  }

  // Locate the laid-out instance (skip table cells / multi-page splits)
  let pageIndex = -1
  let blockSlot = -1
  let existing: LayoutParagraph | undefined
  let occurrences = 0
  for (let pi = 0; pi < layout.pages.length; pi++) {
    const page = layout.pages[pi]!
    for (let bi = 0; bi < page.blocks.length; bi++) {
      const b = page.blocks[bi]!
      if (b.kind === 'paragraph' && b.sectionIndex === sectionIndex && b.blockIndex === blockIndex) {
        occurrences += 1
        pageIndex = pi
        blockSlot = bi
        existing = b
      }
    }
  }
  if (!existing || occurrences !== 1 || pageIndex < 0 || blockSlot < 0) {
    return { layout, needsFullLayout: true }
  }

  const pProps = resolveParagraphProps(doc, block)
  const spacingBefore = twipsToPx(pProps.spacingBefore ?? 0, dpi)
  let indentLeft = twipsToPx(pProps.indentLeft ?? 0, dpi)
  const indentRight = twipsToPx(pProps.indentRight ?? 0, dpi)
  const markers = resolveListMarkers(doc, sectionIndex)
  const marker = markers.get(blockIndex)
  // Must match layoutParagraphBlock: list markers raise the effective left indent.
  // Using props-only indent here double-applies the hanging indent on every keystroke.
  if (marker) {
    indentLeft = Math.max(indentLeft, twipsToPx(marker.indentLeft, dpi))
  }
  const contentX = existing.x - indentLeft
  const contentWidth = existing.width + indentLeft + indentRight
  const cursorY = existing.y - spacingBefore

  const relaid = layoutParagraphBlock(
    doc,
    block,
    sectionIndex,
    blockIndex,
    contentX,
    cursorY,
    contentWidth,
    ctx,
    marker,
  )

  const dy = relaid.height - existing.height
  const page = layout.pages[pageIndex]!
  const contentBottom =
    page.height -
    Math.max(
      twipsToPx(section.properties.margins.bottom, dpi),
      twipsToPx(section.properties.margins.footer, dpi) + 8,
    )

  // Build updated page blocks, shifting followers when height changes
  const newBlocks = page.blocks.map((b, i) => {
    if (i === blockSlot) return relaid
    if (dy !== 0 && i > blockSlot) {
      if (b.kind === 'paragraph') {
        return shiftParagraphY(b, b.y + dy)
      }
      if (b.kind === 'table') {
        return shiftTableLayoutY(b, b.y + dy)
      }
    }
    return b
  })

  const last = newBlocks[newBlocks.length - 1]
  const pageBottom = last ? last.y + last.height : 0
  const overflows = pageBottom > contentBottom + 1
  const underflows = dy < 0 && pageIndex < layout.pages.length - 1

  const newPage: LayoutPage = {
    ...page,
    blocks: newBlocks,
    paragraphs: newBlocks.filter((b): b is LayoutParagraph => b.kind === 'paragraph'),
  }
  const pages = layout.pages.map((p, i) => (i === pageIndex ? newPage : p))
  return {
    layout: { pages },
    needsFullLayout: overflows || underflows || Math.abs(dy) > 0.5,
  }
}

/**
 * Fast path for typing inside a table cell: re-measure a single cell paragraph
 * in place. Follows the same philosophy as {@link patchLayoutParagraph} — paint
 * immediately, and request a full reflow when the paragraph height changes (which
 * can grow the row / shift the rest of the table and pagination).
 */
export function patchLayoutCellParagraph(
  doc: Document,
  layout: LayoutResult,
  sectionIndex: number,
  blockIndex: number,
  cellPath: { row: number; cell: number; para: number },
  options: LayoutOptions = {},
): PatchLayoutResult {
  const dpi = options.dpi ?? 96
  const measurer = options.measurer ?? createApproximateMeasurer()
  const ctx: Ctx = { measurer, dpi }

  const section = doc.sections[sectionIndex]
  const block = section?.blocks[blockIndex]
  if (!block || block.type !== 'table') {
    return { layout, needsFullLayout: true }
  }
  const modelPara = block.rows[cellPath.row]?.cells[cellPath.cell]?.blocks[cellPath.para]
  if (!modelPara || modelPara.type !== 'paragraph') {
    return { layout, needsFullLayout: true }
  }

  // Locate the single laid-out table + cell paragraph. Bail if the table is
  // split across pages / rows in a way that makes the target ambiguous.
  let pageIndex = -1
  let blockSlot = -1
  let cellArrIndex = -1
  let paraArrIndex = -1
  let existing: LayoutParagraph | undefined
  let occurrences = 0
  for (let pi = 0; pi < layout.pages.length; pi++) {
    const page = layout.pages[pi]!
    for (let bi = 0; bi < page.blocks.length; bi++) {
      const b = page.blocks[bi]!
      if (b.kind !== 'table' || b.sectionIndex !== sectionIndex || b.blockIndex !== blockIndex) continue
      for (let ci = 0; ci < b.cells.length; ci++) {
        const c = b.cells[ci]!
        if (c.rowIndex !== cellPath.row || c.cellIndex !== cellPath.cell) continue
        for (let pj = 0; pj < c.paragraphs.length; pj++) {
          const p = c.paragraphs[pj]!
          if (p.cell && p.cell.para === cellPath.para) {
            occurrences += 1
            pageIndex = pi
            blockSlot = bi
            cellArrIndex = ci
            paraArrIndex = pj
            existing = p
          }
        }
      }
    }
  }
  if (!existing || occurrences !== 1) {
    return { layout, needsFullLayout: true }
  }

  const pProps = resolveParagraphProps(doc, modelPara)
  const spacingBefore = twipsToPx(pProps.spacingBefore ?? 0, dpi)
  const indentLeft = twipsToPx(pProps.indentLeft ?? 0, dpi)
  const indentRight = twipsToPx(pProps.indentRight ?? 0, dpi)
  const contentX = existing.x - indentLeft
  const contentWidth = existing.width + indentLeft + indentRight
  const cursorY = existing.y - spacingBefore

  const relaid = layoutParagraphBlock(
    doc,
    modelPara,
    sectionIndex,
    blockIndex,
    contentX,
    cursorY,
    contentWidth,
    ctx,
    undefined,
    { row: cellPath.row, cell: cellPath.cell, para: cellPath.para },
  )

  const dy = relaid.height - existing.height

  const page = layout.pages[pageIndex]!
  const table = page.blocks[blockSlot] as LayoutTable
  const cell = table.cells[cellArrIndex]!
  // Replace the target paragraph; shift later paragraphs in the same cell so the
  // immediate paint stays coherent until the debounced full reflow lands.
  const newParagraphs = cell.paragraphs.map((p, i) => {
    if (i === paraArrIndex) return relaid
    if (dy !== 0 && i > paraArrIndex) return shiftParagraphY(p, p.y + dy)
    return p
  })
  const newCell: LayoutTableCell = { ...cell, paragraphs: newParagraphs }
  const newCells = table.cells.map((c, i) => (i === cellArrIndex ? newCell : c))
  const newTable: LayoutTable = { ...table, cells: newCells }
  const newBlocks = page.blocks.map((b, i) => (i === blockSlot ? newTable : b))
  const newPage: LayoutPage = {
    ...page,
    blocks: newBlocks,
    paragraphs: newBlocks.filter((b): b is LayoutParagraph => b.kind === 'paragraph'),
  }
  const pages = layout.pages.map((p, i) => (i === pageIndex ? newPage : p))

  return {
    layout: { pages },
    // Any height change can grow the row and shift subsequent rows / blocks and
    // pagination, so defer to a full layout; a pure in-cell edit paints instantly.
    needsFullLayout: Math.abs(dy) > 0.5,
  }
}

function shiftParagraphY(para: LayoutParagraph, y: number): LayoutParagraph {
  const dy = y - para.y
  return {
    ...para,
    y,
    lines: para.lines.map((l) => ({ ...l, y: l.y + dy })),
    ...(para.marker ? { marker: { ...para.marker, y: para.marker.y + dy } } : {}),
  }
}

function shiftTableLayoutY(table: LayoutTable, y: number): LayoutTable {
  const dy = y - table.y
  return {
    ...table,
    y,
    cells: table.cells.map((c) => ({
      ...c,
      y: c.y + dy,
      paragraphs: c.paragraphs.map((p) => shiftParagraphY(p, p.y + dy)),
    })),
  }
}
