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
          const table = document.createElement('div')
          table.setAttribute('role', 'grid')
          table.setAttribute('aria-label', 'Table')
          table.setAttribute('aria-rowcount', String(block.rows.length))
          const colCount = Math.max(0, ...block.rows.map((r) => r.cells.length))
          table.setAttribute('aria-colcount', String(colCount))
          for (let ri = 0; ri < block.rows.length; ri++) {
            const row = block.rows[ri]!
            const tr = document.createElement('div')
            tr.setAttribute('role', 'row')
            tr.setAttribute('aria-rowindex', String(ri + 1))
            for (let ci = 0; ci < row.cells.length; ci++) {
              const cell = row.cells[ci]!
              if (cell.props.vMerge === 'continue') continue
              const td = document.createElement('div')
              td.setAttribute('role', 'gridcell')
              td.setAttribute('aria-rowindex', String(ri + 1))
              td.setAttribute('aria-colindex', String(ci + 1))
              if (cell.props.gridSpan && cell.props.gridSpan > 1) {
                td.setAttribute('aria-colspan', String(cell.props.gridSpan))
              }
              td.textContent =
                cell.blocks.map((p) => paragraphPlainText(p)).join(' ') || '\u00a0'
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
