import type { Document } from '@almadocx/core'
import { paragraphPlainText } from '@almadocx/core'

/**
 * Parallel off-screen ARIA document mirror.
 * Canvas is opaque to AT; this DOM tree exposes structure and text.
 */
export class A11yMirror {
  readonly root: HTMLElement

  constructor(parent: HTMLElement) {
    this.root = document.createElement('div')
    this.root.className = 'almadocx-a11y-mirror'
    this.root.setAttribute('role', 'document')
    this.root.setAttribute('aria-label', 'Document')
    Object.assign(this.root.style, {
      position: 'absolute',
      width: '1px',
      height: '1px',
      padding: '0',
      margin: '-1px',
      overflow: 'hidden',
      clip: 'rect(0, 0, 0, 0)',
      whiteSpace: 'nowrap',
      border: '0',
    })
    parent.appendChild(this.root)
  }

  sync(doc: Document, caretOffsetAnnouncement?: string): void {
    this.root.replaceChildren()
    for (const section of doc.sections) {
      const sectionEl = document.createElement('div')
      sectionEl.setAttribute('role', 'region')
      sectionEl.setAttribute('aria-label', 'Section')

      if (section.header) {
        const header = document.createElement('header')
        header.setAttribute('aria-label', 'Header')
        for (const block of section.header.blocks) {
          const p = document.createElement('p')
          p.textContent = paragraphPlainText(block) || '\u00a0'
          header.appendChild(p)
        }
        sectionEl.appendChild(header)
      }

      let listEl: HTMLElement | undefined
      let listKind: 'ul' | 'ol' | undefined

      const flushList = () => {
        if (listEl) sectionEl.appendChild(listEl)
        listEl = undefined
        listKind = undefined
      }

      for (const block of section.blocks) {
        if (block.type === 'table') {
          flushList()
          const table = document.createElement('table')
          table.setAttribute('role', 'table')
          for (const row of block.rows) {
            const tr = document.createElement('tr')
            for (const cell of row.cells) {
              const td = document.createElement(row.props.header ? 'th' : 'td')
              td.textContent = cell.blocks.map((p) => paragraphPlainText(p)).join(' ') || '\u00a0'
              tr.appendChild(td)
            }
            table.appendChild(tr)
          }
          sectionEl.appendChild(table)
          continue
        }

        if (block.props.numPr) {
          const kind = block.props.numPr.numId === '2' ? 'ol' : 'ul'
          if (listKind !== kind) {
            flushList()
            listKind = kind
            listEl = document.createElement(kind)
          }
          const li = document.createElement('li')
          li.textContent = paragraphPlainText(block) || '\u00a0'
          listEl!.appendChild(li)
        } else {
          flushList()
          const p = document.createElement('p')
          p.setAttribute('role', 'paragraph')
          const parts: string[] = []
          for (const run of block.runs) {
            if (run.content.type === 'image') {
              parts.push(run.content.alt ?? 'Image')
            } else if (run.content.type === 'text') {
              parts.push(run.content.text)
            }
          }
          p.textContent = parts.join('') || '\u00a0'
          sectionEl.appendChild(p)
        }
      }
      flushList()

      if (section.footer) {
        const footer = document.createElement('footer')
        footer.setAttribute('aria-label', 'Footer')
        for (const block of section.footer.blocks) {
          const p = document.createElement('p')
          p.textContent = paragraphPlainText(block) || '\u00a0'
          footer.appendChild(p)
        }
        sectionEl.appendChild(footer)
      }

      this.root.appendChild(sectionEl)
    }
    if (caretOffsetAnnouncement) {
      this.root.setAttribute('aria-description', caretOffsetAnnouncement)
    }
  }

  destroy(): void {
    this.root.remove()
  }
}
