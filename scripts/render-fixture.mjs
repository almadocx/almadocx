#!/usr/bin/env node
/**
 * Headless-render a DOCX fixture via @almadocx/core + canvas paint into one PNG per page.
 * Usage: node scripts/render-fixture.mjs [fixture.docx] [outDir]
 */
import { readFileSync, mkdirSync, writeFileSync, unlinkSync } from 'node:fs'
import { dirname, resolve, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import * as esbuild from 'esbuild'
import { chromium } from 'playwright'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const docxPath = resolve(root, process.argv[2] ?? 'fixtures/iou-form.docx')
const outDir = resolve(process.argv[3] ?? '/tmp/iou-compare/alma')
mkdirSync(outDir, { recursive: true })

const entry = join(outDir, 'entry.mjs')
const bundle = join(outDir, 'bundle.js')

writeFileSync(
  entry,
  `
import { loadDocument, layoutDocument } from '${join(root, 'packages/core/src/index.ts')}';
import { paintDocument, createCanvasMeasurer } from '${join(root, 'packages/canvas/src/render/paint.ts')}';

async function loadImages(doc) {
  const images = new Map()
  await Promise.all(
    Object.values(doc.media).map(async (item) => {
      if (!item?.bytes?.byteLength) return
      const blob = new Blob([item.bytes], { type: item.contentType || 'image/png' })
      const url = URL.createObjectURL(blob)
      try {
        const img = new Image()
        img.src = url
        await img.decode()
        images.set(item.id, img)
      } finally {
        URL.revokeObjectURL(url)
      }
    }),
  )
  return images
}

export async function renderPages(bytesArr) {
  const bytes = new Uint8Array(bytesArr)
  const doc = loadDocument(bytes, 'docx')
  const images = await loadImages(doc)
  const measureCanvas = document.createElement('canvas')
  const measureCtx = measureCanvas.getContext('2d')
  const measurer = createCanvasMeasurer(measureCtx)
  const layout = layoutDocument(doc, { measurer })
  const selection = {
    anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 },
    focus: { sectionIndex: 0, blockIndex: 0, offset: 0 },
  }
  const out = []
  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i]
    const single = { pages: [page] }
    const canvas = document.createElement('canvas')
    const ctx = canvas.getContext('2d')
    paintDocument(canvas, ctx, {
      layout: single,
      doc,
      selection,
      zoom: 1,
      dpr: 1,
      caretVisible: false,
      virtualized: false,
      pageChrome: false,
      images,
    })
    // Crop to the paper rect (PAGE_TOP_PAD above; may have side margin if canvas wider).
    const pageW = page.width
    const pageH = page.height
    const topPad = 24
    const left = Math.max(0, Math.floor((canvas.width - pageW) / 2))
    const crop = document.createElement('canvas')
    crop.width = pageW
    crop.height = pageH
    const cctx = crop.getContext('2d')
    cctx.fillStyle = '#ffffff'
    cctx.fillRect(0, 0, pageW, pageH)
    cctx.drawImage(canvas, left, topPad, pageW, pageH, 0, 0, pageW, pageH)
    out.push({
      index: i + 1,
      width: pageW,
      height: pageH,
      dataUrl: crop.toDataURL('image/png'),
      blockCount: (page.blocks ?? page.paragraphs)?.length ?? 0,
      footerParas: page.footer?.paragraphs?.length ?? 0,
    })
  }
  return {
    pageCount: layout.pages.length,
    pages: out.map(({ dataUrl, ...meta }) => meta),
    pngs: out.map((p) => p.dataUrl),
  }
}
`,
)

await esbuild.build({
  entryPoints: [entry],
  bundle: true,
  outfile: bundle,
  format: 'esm',
  platform: 'browser',
  target: 'es2022',
  loader: { '.ts': 'ts' },
  define: { 'process.env.NODE_ENV': '"production"' },
})

const bytes = [...readFileSync(docxPath)]
writeFileSync(
  join(outDir, 'index.html'),
  `<!doctype html>
<html><body style="margin:0">
<script type="module">
  import { renderPages } from './bundle.js'
  const bytes = ${JSON.stringify(bytes)}
  window.__done = renderPages(bytes).then((r) => { window.__result = r; return r })
</script>
</body></html>`,
)

const browser = await chromium.launch({
  executablePath: process.env.CHROMIUM_PATH || '/usr/bin/chromium',
  args: ['--no-sandbox', '--disable-gpu', '--allow-file-access-from-files'],
})
const page = await browser.newPage({ viewport: { width: 1000, height: 1200 }, deviceScaleFactor: 1 })
await page.goto(`file://${join(outDir, 'index.html')}`, { waitUntil: 'networkidle' })
const result = await page.evaluate(async () => {
  const r = await window.__done
  return r
})
console.log('meta', JSON.stringify({ pageCount: result.pageCount, pages: result.pages }))

for (let i = 0; i < result.pngs.length; i++) {
  const b64 = result.pngs[i].replace(/^data:image\/png;base64,/, '')
  const file = join(outDir, `page-${i + 1}.png`)
  writeFileSync(file, Buffer.from(b64, 'base64'))
  console.log('wrote', file)
}
await browser.close()
try {
  unlinkSync(entry)
} catch {
  /* ignore */
}
