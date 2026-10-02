/**
 * HTML clipboard sanitization for paste from Word / Google Docs / LibreOffice / browsers.
 * Never evaluates scripts; strips dangerous tags/attrs; maps a safe subset to CharacterProps.
 */

import {
  createEmptyParagraph,
  createTextRun,
  nextId,
  type CharacterProps,
  type Paragraph,
  type Run,
} from '@almadocx/core'

const ALLOWED_TAGS = new Set([
  'B',
  'STRONG',
  'I',
  'EM',
  'U',
  'S',
  'STRIKE',
  'DEL',
  'SPAN',
  'P',
  'DIV',
  'BR',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'UL',
  'OL',
  'LI',
  'SUB',
  'SUP',
  'FONT',
])

const DANGEROUS_ATTR = /^(on|formaction|xlink:href)/i

export interface PastedFragment {
  paragraphs: Paragraph[]
}

function parseColor(value: string | null | undefined): string | undefined {
  if (!value) return undefined
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value.trim())
  if (hex) {
    const h = hex[1]!
    if (h.length === 3) {
      return `#${h[0]}${h[0]}${h[1]}${h[1]}${h[2]}${h[2]}`.toLowerCase()
    }
    return `#${h.toLowerCase()}`
  }
  const rgb = /^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i.exec(value.trim())
  if (rgb) {
    const r = Number(rgb[1]).toString(16).padStart(2, '0')
    const g = Number(rgb[2]).toString(16).padStart(2, '0')
    const b = Number(rgb[3]).toString(16).padStart(2, '0')
    return `#${r}${g}${b}`
  }
  return undefined
}

function styleToProps(style: string | null, base: CharacterProps): CharacterProps {
  const props = { ...base }
  if (!style) return props
  for (const part of style.split(';')) {
    const [rawK, rawV] = part.split(':')
    if (!rawK || rawV === undefined) continue
    const k = rawK.trim().toLowerCase()
    const v = rawV.trim()
    if (k === 'font-weight' && (v === 'bold' || Number(v) >= 600)) props.bold = true
    if (k === 'font-style' && v === 'italic') props.italic = true
    if (k === 'text-decoration' && v.includes('underline')) props.underline = true
    if (k === 'text-decoration' && (v.includes('line-through') || v.includes('strike'))) {
      props.strike = true
    }
    if (k === 'font-size') {
      const pt = Number.parseFloat(v)
      if (!Number.isNaN(pt)) {
        props.fontSizePt = v.endsWith('px') ? pt * 0.75 : pt
      }
    }
    if (k === 'font-family') {
      const family = v.split(',')[0]?.replace(/['"]/g, '').trim()
      if (family) props.fontFamily = family
    }
    if (k === 'color') {
      const c = parseColor(v)
      if (c) props.color = c
    }
    if (k === 'background-color' || k === 'background') {
      const c = parseColor(v)
      if (c) props.highlight = c
    }
  }
  return props
}

function tagProps(tag: string, el: Element, base: CharacterProps): CharacterProps {
  let props = { ...base }
  const name = tag.toUpperCase()
  if (name === 'B' || name === 'STRONG') props.bold = true
  if (name === 'I' || name === 'EM') props.italic = true
  if (name === 'U') props.underline = true
  if (name === 'S' || name === 'STRIKE' || name === 'DEL') props.strike = true
  if (name === 'SUB') props.verticalAlign = 'subscript'
  if (name === 'SUP') props.verticalAlign = 'superscript'
  if (name === 'FONT') {
    const face = el.getAttribute('face')
    const color = el.getAttribute('color')
    const size = el.getAttribute('size')
    if (face) props.fontFamily = face
    const c = parseColor(color)
    if (c) props.color = c
    if (size) {
      const n = Number(size)
      if (!Number.isNaN(n)) props.fontSizePt = 7 + n * 2
    }
  }
  props = styleToProps(el.getAttribute('style'), props)
  // Google Docs / Word often put bold via class — ignore classes (untrusted)
  return props
}

function sanitizeNode(node: Node, into: Element): void {
  if (node.nodeType === Node.TEXT_NODE) {
    into.appendChild(document.createTextNode(node.textContent!))
    return
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return
  const el = node as Element
  const tag = el.tagName.toUpperCase()
  if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'IFRAME' || tag === 'OBJECT' || tag === 'EMBED') {
    return
  }
  if (!ALLOWED_TAGS.has(tag)) {
    // Unwrap unknown tags but keep children
    for (const child of Array.from(el.childNodes)) sanitizeNode(child, into)
    return
  }
  const clean = document.createElement(tag)
  for (const attr of Array.from(el.attributes)) {
    const name = attr.name.toLowerCase()
    if (DANGEROUS_ATTR.test(name) || name === 'href' || name === 'src') continue
    if (name === 'style' || name === 'face' || name === 'color' || name === 'size') {
      clean.setAttribute(attr.name, attr.value)
    }
  }
  for (const child of Array.from(el.childNodes)) sanitizeNode(child, clean)
  into.appendChild(clean)
}

export function sanitizeHtml(html: string): string {
  if (typeof document === 'undefined') {
    // Node fallback: strip tags crudely
    return html
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/on\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
  }
  const template = document.createElement('template')
  // Avoid script execution: assign to innerHTML of template content via DOMParser when available
  const parsed = new DOMParser().parseFromString(`<div id="root">${html}</div>`, 'text/html')
  const root = parsed.getElementById('root')
  if (!root) return ''
  const out = document.createElement('div')
  for (const child of Array.from(root.childNodes)) sanitizeNode(child, out)
  void template
  return out.innerHTML
}

function walk(
  node: Node,
  props: CharacterProps,
  runs: Paragraph['runs'],
  paragraphs: Paragraph[],
  current: { p: Paragraph },
): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent!
    if (text) current.p.runs.push(createTextRun(text, props))
    return
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return
  const el = node as Element
  const tag = el.tagName.toUpperCase()
  if (tag === 'BR') {
    current.p.runs.push({
      id: nextId('r'),
      props,
      content: { type: 'break', breakType: 'line' },
    })
    return
  }
  if (tag === 'P' || tag === 'DIV' || tag === 'LI' || /^H[1-6]$/.test(tag)) {
    if (current.p.runs.length > 0 || paragraphs.length === 0) {
      // flush previous if it has content
      if (current.p.runs.length > 0) paragraphs.push(current.p)
      current.p = createEmptyParagraph()
      current.p.runs = []
    }
    const nextProps = tagProps(tag, el, props)
    if (/^H[1-6]$/.test(tag)) {
      nextProps.bold = true
      const fontSizePt = 24 - (Number(tag[1]) - 1) * 2
      nextProps.fontSizePt = fontSizePt
      current.p = createEmptyParagraph(`Heading${tag[1]!}`)
      current.p.runs = []
    }
    for (const child of Array.from(el.childNodes)) {
      walk(child, nextProps, runs, paragraphs, current)
    }
    if (current.p.runs.length === 0) current.p.runs.push(createTextRun(''))
    // Ensure heading character props landed on runs
    if (/^H[1-6]$/.test(tag)) {
      const fontSizePt = nextProps.fontSizePt!
      current.p.runs = current.p.runs.map((r) => ({
        ...r,
        props: {
          ...nextProps,
          ...r.props,
          bold: true,
          fontSizePt,
        },
      }))
    }
    if (tag === 'LI') {
      current.p.props = { ...current.p.props, numPr: { numId: '1', ilvl: 0 } }
    }
    paragraphs.push(current.p)
    current.p = createEmptyParagraph()
    current.p.runs = []
    return
  }
  const nextProps = tagProps(tag, el, props)
  for (const child of Array.from(el.childNodes)) {
    walk(child, nextProps, runs, paragraphs, current)
  }
}

/**
 * Parse sanitized HTML into paragraph fragments for paste.
 */
export function htmlToFragment(html: string): PastedFragment {
  const safe = sanitizeHtml(html)
  if (typeof document === 'undefined') {
    // Plain fallback
    const lines = safe.replace(/<[^>]+>/g, '').split(/\n/)
    return {
      paragraphs: lines.map((line) => {
        const p = createEmptyParagraph()
        p.runs = [createTextRun(line)]
        return p
      }),
    }
  }
  const parsed = new DOMParser().parseFromString(`<div id="root">${safe}</div>`, 'text/html')
  const root = parsed.getElementById('root')
  const paragraphs: Paragraph[] = []
  const current = { p: createEmptyParagraph() }
  current.p.runs = []
  if (root) {
    for (const child of Array.from(root.childNodes)) {
      walk(child, {}, [], paragraphs, current)
    }
  }
  if (current.p.runs.length > 0) paragraphs.push(current.p)
  if (paragraphs.length === 0) {
    paragraphs.push(createEmptyParagraph())
  }
  return { paragraphs }
}

export function plainTextToFragment(text: string): PastedFragment {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n')
  return {
    paragraphs: lines.map((line) => {
      const p = createEmptyParagraph()
      p.runs = [createTextRun(line)]
      return p
    }),
  }
}

/** Build a minimal HTML snapshot for copy (marks only). */
export function runsToHtml(paragraphs: Paragraph[]): string {
  return paragraphs
    .map((p) => {
      const inner = p.runs
        .map((r: Run) => {
          if (r.content.type !== 'text') {
            if (r.content.type === 'tab') return '&emsp;'
            if (r.content.type === 'break') return '<br/>'
            return ''
          }
          let t = r.content.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
          if (r.props.bold) t = `<strong>${t}</strong>`
          if (r.props.italic) t = `<em>${t}</em>`
          if (r.props.underline) t = `<u>${t}</u>`
          if (r.props.strike) t = `<s>${t}</s>`
          return t
        })
        .join('')
      return `<p>${inner || '<br/>'}</p>`
    })
    .join('')
}
