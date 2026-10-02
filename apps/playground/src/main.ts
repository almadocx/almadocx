import { createEmptyDocument, applyOp } from '@almadocx/core'
import { mountEditor, type EditorHandle } from '@almadocx/canvas'
import './styles.css'

function sampleDocument() {
  let doc = createEmptyDocument('docx')
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 0, offset: 0 },
    text: 'Welcome to Almadocx',
    props: { bold: true, fontSizePt: 22, fontFamily: 'Fraunces' },
  }).doc
  doc = applyOp(doc, {
    type: 'splitParagraph',
    position: { sectionIndex: 0, blockIndex: 0, offset: paragraphLen(doc, 0) },
  }).doc
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 1, offset: 0 },
    text:
      'A canvas-based document engine with DOCX and ODT round-trip, invertible editing operations, and a fidelity-first layout pipeline. Start typing, or open a file.',
  }).doc
  doc = applyOp(doc, {
    type: 'splitParagraph',
    position: { sectionIndex: 0, blockIndex: 1, offset: paragraphLen(doc, 1) },
  }).doc
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 2, offset: 0 },
    text: 'Bullet item one',
  }).doc
  doc = applyOp(doc, {
    type: 'setList',
    sectionIndex: 0,
    blockIndex: 2,
    numPr: { numId: '1', ilvl: 0 },
  }).doc
  doc = applyOp(doc, {
    type: 'splitParagraph',
    position: { sectionIndex: 0, blockIndex: 2, offset: paragraphLen(doc, 2) },
  }).doc
  doc = applyOp(doc, {
    type: 'insertText',
    position: { sectionIndex: 0, blockIndex: 3, offset: 0 },
    text: 'Bullet item two',
  }).doc
  doc = applyOp(doc, {
    type: 'setList',
    sectionIndex: 0,
    blockIndex: 3,
    numPr: { numId: '1', ilvl: 0 },
  }).doc
  doc.sections[0]!.header = {
    id: 'hdr_1',
    blocks: [
      {
        id: 'hp_1',
        type: 'paragraph',
        props: { alignment: 'right' },
        runs: [{ id: 'hr_1', props: { fontSizePt: 9, color: '#5c534a' }, content: { type: 'text', text: 'Almadocx · {{PAGE}}' } }],
      },
    ],
  }
  return doc
}

function paragraphLen(doc: ReturnType<typeof createEmptyDocument>, blockIndex: number): number {
  const block = doc.sections[0]?.blocks[blockIndex]
  if (!block || block.type !== 'paragraph') return 0
  return block.runs.reduce((n, r) => {
    if (r.content.type === 'text') return n + r.content.text.length
    return n + 1
  }, 0)
}

const app = document.querySelector('#app')
if (!app) throw new Error('#app missing')

app.innerHTML = `
  <header class="topbar">
    <div class="brand">Almadocx<span>Canvas document editor · DOCX &amp; ODT</span></div>
  </header>
  <div class="toolbar" role="toolbar" aria-label="Document tools">
    <div class="group">
      <label class="label-btn">Open<input type="file" id="open" accept=".docx,.odt,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.oasis.opendocument.text" /></label>
      <button type="button" id="save-docx" class="primary">Save DOCX</button>
      <button type="button" id="save-odt">Save ODT</button>
    </div>
    <div class="group">
      <button type="button" id="undo" title="Ctrl+Z">Undo</button>
      <button type="button" id="redo" title="Ctrl+Y">Redo</button>
    </div>
    <div class="group">
      <button type="button" id="bold" title="Ctrl+B"><strong>B</strong></button>
      <button type="button" id="italic" title="Ctrl+I"><em>I</em></button>
      <button type="button" id="underline" title="Ctrl+U"><span style="text-decoration:underline">U</span></button>
      <button type="button" id="strike" title="Strikethrough"><span style="text-decoration:line-through">S</span></button>
    </div>
    <div class="group">
      <button type="button" id="align-left" title="Align left">⟸</button>
      <button type="button" id="align-center" title="Align center">⇔</button>
      <button type="button" id="align-right" title="Align right">⟹</button>
      <button type="button" id="align-justify" title="Justify">☰</button>
    </div>
    <div class="group">
      <button type="button" id="bullet" title="Bullet list">• List</button>
      <button type="button" id="number" title="Numbered list">1. List</button>
      <button type="button" id="clear-list" title="Clear list">Clear list</button>
    </div>
    <div class="group find-group">
      <input type="search" id="find-query" placeholder="Find…" aria-label="Find" />
      <input type="search" id="replace-query" placeholder="Replace…" aria-label="Replace" />
      <button type="button" id="find-next">Find</button>
      <button type="button" id="replace-one">Replace</button>
      <button type="button" id="replace-all">All</button>
    </div>
    <div class="group">
      <button type="button" id="zoom-out" aria-label="Zoom out">−</button>
      <span class="zoom-label" id="zoom-label">100%</span>
      <button type="button" id="zoom-in" aria-label="Zoom in">+</button>
      <button type="button" id="zoom-fit-width" title="Fit page width">Fit width</button>
      <button type="button" id="zoom-fit-page" title="Fit whole page">Fit page</button>
      <button type="button" id="zoom-100" title="100%">100%</button>
      <button type="button" id="zoom-two" title="Two pages">2 pages</button>
    </div>
  </div>
  <div class="editor-shell"><div id="editor" class="almadocx-host" aria-label="Editor"></div></div>
  <div class="toast" id="toast" role="status" aria-live="polite"></div>
`

const editorHost = document.querySelector('#editor') as HTMLElement
const toastEl = document.querySelector('#toast') as HTMLElement

function toast(message: string, kind: 'info' | 'error' = 'info') {
  toastEl.textContent = message
  toastEl.classList.toggle('error', kind === 'error')
  toastEl.classList.add('show')
  window.setTimeout(() => toastEl.classList.remove('show'), 3200)
}

const editor: EditorHandle = mountEditor(editorHost, {
  document: sampleDocument(),
  zoom: 1,
})

const zoomLabel = document.querySelector('#zoom-label') as HTMLElement

document.querySelector('#open')?.addEventListener('change', (ev) => {
  void (async () => {
    const input = ev.target as HTMLInputElement
    const file = input.files?.[0]
    if (!file) return
    try {
      const buf = new Uint8Array(await file.arrayBuffer())
      const name = file.name.toLowerCase()
      const format = name.endsWith('.odt') ? 'odt' : name.endsWith('.docx') ? 'docx' : undefined
      editor.loadBytes(buf, format)
      toast(`Opened ${file.name}`)
    } catch (e) {
      toast(e instanceof Error ? e.message : 'Failed to open file', 'error')
    } finally {
      input.value = ''
    }
  })()
})

async function download(filename: string, bytes: Uint8Array) {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  const blob = new Blob([copy], { type: 'application/octet-stream' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

document.querySelector('#save-docx')?.addEventListener('click', () => {
  try {
    void download('document.docx', editor.saveBytes('docx'))
    toast('Saved DOCX')
  } catch (e) {
    toast(e instanceof Error ? e.message : 'Save failed', 'error')
  }
})

document.querySelector('#save-odt')?.addEventListener('click', () => {
  try {
    void download('document.odt', editor.saveBytes('odt'))
    toast('Saved ODT')
  } catch (e) {
    toast(e instanceof Error ? e.message : 'Save failed', 'error')
  }
})

document.querySelector('#undo')?.addEventListener('click', () => editor.undo())
document.querySelector('#redo')?.addEventListener('click', () => editor.redo())
document.querySelector('#bold')?.addEventListener('click', () => editor.bold())
document.querySelector('#italic')?.addEventListener('click', () => editor.italic())
document.querySelector('#underline')?.addEventListener('click', () => editor.underline())
document.querySelector('#strike')?.addEventListener('click', () => editor.strike())
document.querySelector('#align-left')?.addEventListener('click', () => editor.align('left'))
document.querySelector('#align-center')?.addEventListener('click', () => editor.align('center'))
document.querySelector('#align-right')?.addEventListener('click', () => editor.align('right'))
document.querySelector('#align-justify')?.addEventListener('click', () => editor.align('justify'))
document.querySelector('#bullet')?.addEventListener('click', () => editor.bulletList())
document.querySelector('#number')?.addEventListener('click', () => editor.numberedList())
document.querySelector('#clear-list')?.addEventListener('click', () => editor.clearList())

const findInput = document.querySelector('#find-query') as HTMLInputElement
const replaceInput = document.querySelector('#replace-query') as HTMLInputElement
document.querySelector('#find-next')?.addEventListener('click', () => {
  const q = findInput.value
  if (!q) return
  const m = editor.findNext({ query: q })
  toast(m ? 'Match selected' : 'No matches')
})
document.querySelector('#replace-one')?.addEventListener('click', () => {
  const q = findInput.value
  if (!q) return
  const ok = editor.replace({ query: q }, replaceInput.value)
  toast(ok ? 'Replaced' : 'No match to replace')
})
document.querySelector('#replace-all')?.addEventListener('click', () => {
  const q = findInput.value
  if (!q) return
  const n = editor.replaceAll({ query: q }, replaceInput.value)
  toast(`Replaced ${String(n)} occurrence${n === 1 ? '' : 's'}`)
})

document.querySelector('#zoom-in')?.addEventListener('click', () => {
  editor.setZoom(editor.getZoom() + 0.1)
  zoomLabel.textContent = `${Math.round(editor.getZoom() * 100)}%`
})
document.querySelector('#zoom-out')?.addEventListener('click', () => {
  editor.setZoom(editor.getZoom() - 0.1)
  zoomLabel.textContent = `${Math.round(editor.getZoom() * 100)}%`
})
const syncZoomLabel = () => {
  const mode = editor.getZoomMode()
  if (mode === 'fitWidth') zoomLabel.textContent = 'Fit W'
  else if (mode === 'fitPage') zoomLabel.textContent = 'Fit'
  else if (mode === 'twoPages') zoomLabel.textContent = '2×'
  else zoomLabel.textContent = `${Math.round(editor.getZoom() * 100)}%`
}
document.querySelector('#zoom-fit-width')?.addEventListener('click', () => {
  editor.setZoomMode('fitWidth')
  syncZoomLabel()
})
document.querySelector('#zoom-fit-page')?.addEventListener('click', () => {
  editor.setZoomMode('fitPage')
  syncZoomLabel()
})
document.querySelector('#zoom-100')?.addEventListener('click', () => {
  editor.setZoom(1)
  syncZoomLabel()
})
document.querySelector('#zoom-two')?.addEventListener('click', () => {
  editor.setZoomMode('twoPages')
  syncZoomLabel()
})

editor.focus()
