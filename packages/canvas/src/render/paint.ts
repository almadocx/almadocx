import type { DocRange, Document, LayoutPage, LayoutParagraph, LayoutResult } from '@almadocx/core'
import { getCaretScreenRect, offsetToX, paintSelection } from '../selection/selectionPaint.js'
import { pageEditableParagraphs } from '../selection/hitTest.js'
import {
  computePageLayoutMetrics,
  pageOriginInView,
  PAGE_GAP_DEFAULT,
} from './pageLayout.js'

export interface PaintOptions {
  layout: LayoutResult
  doc: Document
  selection: DocRange
  zoom: number
  dpr: number
  pageGap?: number
  caretVisible?: boolean
  images?: Map<string, CanvasImageSource>
  /** Host viewport for virtualized painting (avoids huge canvas buffers). */
  viewport?: { scrollTop: number; scrollLeft: number; width: number; height: number }
  /** When set, canvas is viewport-sized; spacer element should match content size. */
  virtualized?: boolean
  /** Side-by-side page columns (Word "Multiple pages"). */
  pageColumns?: number
  /** Draw drop shadow + page edge stroke (editor chrome). Default true. */
  pageChrome?: boolean
}

export interface ContentSize {
  width: number
  height: number
}

export function measureContentSize(
  layout: LayoutResult,
  zoom: number,
  pageGap = PAGE_GAP_DEFAULT,
  pageColumns = 1,
): ContentSize {
  const m = computePageLayoutMetrics(layout, zoom, pageGap, pageColumns)
  return { width: m.contentWidth, height: m.contentHeight }
}

export function paintDocument(
  canvas: HTMLCanvasElement,
  ctx: CanvasRenderingContext2D,
  options: PaintOptions,
): ContentSize {
  const { layout, selection, zoom, dpr, caretVisible = true, doc, images } = options
  const pageGap = options.pageGap ?? PAGE_GAP_DEFAULT
  const pageColumns = Math.max(1, options.pageColumns ?? 1)
  const vp = options.viewport
  const metrics = computePageLayoutMetrics(layout, zoom, pageGap, pageColumns)
  const viewportWidth = vp?.width ?? metrics.contentWidth
  const totalHeight = metrics.contentHeight

  const virtualized = Boolean(options.virtualized && vp)

  if (virtualized && vp) {
    const cssW = Math.max(1, Math.floor(vp.width))
    const cssH = Math.max(1, Math.floor(vp.height))
    canvas.width = Math.ceil(cssW * dpr)
    canvas.height = Math.ceil(cssH * dpr)
    canvas.style.width = `${cssW}px`
    canvas.style.height = `${cssH}px`
    canvas.style.position = 'sticky'
    canvas.style.top = '0'
    canvas.style.left = '0'
    canvas.style.right = '0'
    canvas.style.margin = '0 auto'
    canvas.style.zIndex = '1'
    canvas.style.display = 'block'

    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, cssW, cssH)
    ctx.fillStyle = '#e8e4dc'
    ctx.fillRect(0, 0, cssW, cssH)

    ctx.save()
    // Only pan with scroll; horizontal centering is baked into page origins
    ctx.translate(-vp.scrollLeft, -vp.scrollTop)
    paintPages(ctx, layout, selection, zoom, pageGap, metrics, viewportWidth, {
      caretVisible,
      pageChrome: options.pageChrome !== false,
      ...(images ? { images } : {}),
      clipTop: vp.scrollTop - pageGap * zoom,
      clipBottom: vp.scrollTop + vp.height + pageGap * zoom,
    })
    ctx.restore()
  } else {
    // Non-virtualized: center in at least content width (caller may pass viewport via options)
    const frameW = Math.max(metrics.contentWidth, viewportWidth)
    const maxDim = 8192
    const scale = Math.min(1, maxDim / (frameW * dpr), maxDim / (totalHeight * dpr))
    const effectiveDpr = dpr * scale
    canvas.width = Math.max(1, Math.ceil(frameW * effectiveDpr))
    canvas.height = Math.max(1, Math.ceil(totalHeight * effectiveDpr))
    canvas.style.width = `${Math.ceil(frameW)}px`
    canvas.style.height = `${Math.ceil(totalHeight)}px`
    canvas.style.position = ''
    canvas.style.top = ''
    canvas.style.left = ''
    canvas.style.right = ''
    canvas.style.margin = ''

    ctx.setTransform(effectiveDpr, 0, 0, effectiveDpr, 0, 0)
    ctx.clearRect(0, 0, frameW, totalHeight)
    ctx.fillStyle = '#e8e4dc'
    ctx.fillRect(0, 0, frameW, totalHeight)
    paintPages(ctx, layout, selection, zoom, pageGap, metrics, frameW, {
      caretVisible,
      pageChrome: options.pageChrome !== false,
      ...(images ? { images } : {}),
    })
  }

  void doc
  return { width: metrics.contentWidth, height: totalHeight }
}

function paintPages(
  ctx: CanvasRenderingContext2D,
  layout: LayoutResult,
  selection: DocRange,
  zoom: number,
  pageGap: number,
  metrics: ReturnType<typeof computePageLayoutMetrics>,
  viewportWidth: number,
  opts: {
    caretVisible: boolean
    pageChrome?: boolean
    images?: Map<string, CanvasImageSource>
    clipTop?: number
    clipBottom?: number
  },
): void {
  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]!
    const { left: pageLeft, top: pageTop } = pageOriginInView(
      metrics,
      i,
      zoom,
      pageGap,
      viewportWidth,
    )
    const pageBottom = pageTop + page.height * zoom

    const visible =
      opts.clipTop === undefined ||
      opts.clipBottom === undefined ||
      (pageBottom >= opts.clipTop && pageTop <= opts.clipBottom)

    if (!visible) continue

    const chrome = opts.pageChrome !== false
    if (chrome) {
      ctx.fillStyle = 'rgba(40, 32, 24, 0.10)'
      ctx.fillRect(pageLeft + 2 * zoom, pageTop + 2 * zoom, page.width * zoom, page.height * zoom)
    }
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(pageLeft, pageTop, page.width * zoom, page.height * zoom)
    if (chrome) {
      ctx.strokeStyle = 'rgba(40, 32, 24, 0.08)'
      ctx.lineWidth = 1
      ctx.strokeRect(pageLeft + 0.5, pageTop + 0.5, page.width * zoom - 1, page.height * zoom - 1)
    }

    ctx.save()
    ctx.translate(pageLeft, pageTop)
    ctx.scale(zoom, zoom)
    ctx.beginPath()
    ctx.rect(0, 0, page.width, page.height)
    ctx.clip()

    paintHeaderFooter(ctx, page, 'header')

    // Table chrome first so selection highlight paints above cell fills
    for (const block of page.blocks ?? page.paragraphs) {
      if (block.kind !== 'table') continue
      for (const cell of block.cells) {
        if (cell.shading) {
          ctx.fillStyle = cell.shading
          ctx.fillRect(cell.x, cell.y, cell.width, cell.height)
        }
        ctx.strokeStyle = cell.borderColor
        ctx.lineWidth = 1
        ctx.strokeRect(cell.x + 0.5, cell.y + 0.5, cell.width - 1, cell.height - 1)
      }
    }

    paintSelection(ctx, page, selection)

    for (const block of page.blocks ?? page.paragraphs) {
      if (block.kind === 'paragraph') {
        paintParagraph(ctx, block, opts.images)
      } else if (block.kind === 'table') {
        for (const cell of block.cells) {
          for (const p of cell.paragraphs) paintParagraph(ctx, p, opts.images)
        }
      }
    }

    if (opts.caretVisible) paintCaret(ctx, page, selection)
    paintHeaderFooter(ctx, page, 'footer')
    ctx.restore()
  }
}

function paintHeaderFooter(
  ctx: CanvasRenderingContext2D,
  page: LayoutPage,
  which: 'header' | 'footer',
): void {
  const hf = which === 'header' ? page.header : page.footer
  if (!hf) return
  for (const p of hf.paragraphs) paintParagraph(ctx, p)
  if (hf.pageNumberText && hf.paragraphs.length === 0) {
    ctx.font = 'normal 400 11px "Source Sans 3", sans-serif'
    ctx.fillStyle = '#5c534a'
    ctx.textBaseline = 'alphabetic'
    const y = which === 'header' ? 28 : page.height - 28
    const m = ctx.measureText(hf.pageNumberText)
    ctx.fillText(hf.pageNumberText, (page.width - m.width) / 2, y)
  }
}

function paintParagraph(
  ctx: CanvasRenderingContext2D,
  para: LayoutParagraph,
  images?: Map<string, CanvasImageSource>,
): void {
  if (para.marker) {
    ctx.font = para.marker.font
    ctx.fillStyle = para.marker.color
    ctx.textBaseline = 'alphabetic'
    const first = para.lines[0]
    const baseline = first ? first.y + first.baseline : para.y + 12
    ctx.fillText(para.marker.text, para.marker.x, baseline)
  }
  for (const line of para.lines) {
    // Merge consecutive underlined runs so form fill-in lines don't show hairline gaps.
    let ulStart: number | null = null
    let ulEnd = 0
    let ulColor = '#000000'
    const flushUnderline = () => {
      if (ulStart === null || ulEnd <= ulStart) {
        ulStart = null
        return
      }
      ctx.strokeStyle = ulColor
      ctx.lineWidth = 1
      ctx.beginPath()
      const uy = line.y + line.baseline + 1.5
      ctx.moveTo(ulStart, uy)
      ctx.lineTo(ulEnd, uy)
      ctx.stroke()
      ulStart = null
    }
    for (const run of line.runs) {
      if (run.image) {
        flushUnderline()
        const src = images?.get(run.image.mediaId)
        const y = line.y + line.baseline - run.image.height * 0.85
        if (src) {
          ctx.drawImage(src, run.x, y, run.image.width, run.image.height)
        } else {
          ctx.fillStyle = '#d9d2c5'
          ctx.fillRect(run.x, y, run.image.width, run.image.height)
          ctx.strokeStyle = '#a39a8c'
          ctx.strokeRect(run.x + 0.5, y + 0.5, run.image.width - 1, run.image.height - 1)
        }
        continue
      }
      ctx.font = run.font
      ctx.fillStyle = run.color
      ctx.textBaseline = 'alphabetic'
      const x = run.x
      const y = line.y + line.baseline
      if (run.text && run.text !== '\t') ctx.fillText(run.text, x, y)
      if (run.underline && run.width > 0) {
        if (ulStart === null) {
          ulStart = x
          ulColor = run.color
        }
        ulEnd = x + run.width
      } else {
        flushUnderline()
      }
      if (run.strike) {
        ctx.strokeStyle = run.color
        ctx.lineWidth = 1
        ctx.beginPath()
        const mid = line.y + line.height / 2
        ctx.moveTo(x, mid)
        ctx.lineTo(x + run.width, mid)
        ctx.stroke()
      }
      if (run.tabLeader && run.tabLeader !== 'none') {
        ctx.strokeStyle = run.color
        ctx.lineWidth = 1
        ctx.beginPath()
        const yDot = line.y + line.baseline + 1
        if (run.tabLeader === 'dot') {
          for (let dx = x + 2; dx < x + run.width - 2; dx += 4) {
            ctx.fillRect(dx, yDot, 1.5, 1.5)
          }
        } else if (run.tabLeader === 'dash') {
          ctx.setLineDash([4, 4])
          ctx.moveTo(x + 2, yDot)
          ctx.lineTo(x + run.width - 2, yDot)
          ctx.stroke()
          ctx.setLineDash([])
        } else {
          ctx.moveTo(x + 2, yDot)
          ctx.lineTo(x + run.width - 2, yDot)
          ctx.stroke()
        }
      }
    }
    flushUnderline()
  }
}

function paintCaret(ctx: CanvasRenderingContext2D, page: LayoutPage, selection: DocRange): void {
  const pos = selection.focus
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
      const x = offsetToX(line, pos.offset)
      ctx.fillStyle = '#1a1510'
      ctx.fillRect(x, line.y, 1.25, line.height)
      return
    }
  }
}

export function createCanvasMeasurer(ctx: CanvasRenderingContext2D) {
  const cache = new Map<string, { width: number; ascent: number; descent: number }>()
  return {
    measure(text: string, font: string) {
      const key = `${font}\0${text}`
      const hit = cache.get(key)
      if (hit) return hit
      ctx.font = font
      const m = ctx.measureText(text)
      const result = {
        width: m.width,
        ascent: m.actualBoundingBoxAscent || 12,
        descent: m.actualBoundingBoxDescent || 4,
      }
      if (cache.size > 12_000) cache.clear()
      cache.set(key, result)
      return result
    },
  }
}

export { getCaretScreenRect }
