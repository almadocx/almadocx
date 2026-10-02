export type ImeHandler = {
  onCommit: (text: string) => void
  onCompositionUpdate: (text: string) => void
  onCompositionEnd: (text: string) => void
  onKeyDown?: (e: KeyboardEvent) => void
}

/**
 * Hidden textarea input proxy with full IME composition lifecycle.
 * Keyboard shortcuts are handled via keydown on the textarea (where focus lives).
 */
export class InputProxy {
  readonly el: HTMLTextAreaElement
  private composing = false
  private handlers: ImeHandler

  constructor(parent: HTMLElement, handlers: ImeHandler) {
    this.handlers = handlers
    this.el = document.createElement('textarea')
    this.el.setAttribute('aria-hidden', 'true')
    this.el.setAttribute('autocomplete', 'off')
    this.el.tabIndex = 0
    this.el.autocomplete = 'off'
    this.el.autocapitalize = 'off'
    this.el.spellcheck = false
    Object.assign(this.el.style, {
      position: 'absolute',
      left: '0',
      top: '0',
      width: '1px',
      height: '1px',
      opacity: '0',
      resize: 'none',
      overflow: 'hidden',
      zIndex: '0',
      pointerEvents: 'none',
    })
    parent.appendChild(this.el)

    this.el.addEventListener('beforeinput', this.onBeforeInput)
    this.el.addEventListener('compositionstart', this.onCompStart)
    this.el.addEventListener('compositionupdate', this.onCompUpdate)
    this.el.addEventListener('compositionend', this.onCompEnd)
    this.el.addEventListener('input', this.onInput)
    this.el.addEventListener('keydown', this.onKeyDown)
  }

  focus(): void {
    this.el.focus({ preventScroll: true })
  }

  isComposing(): boolean {
    return this.composing
  }

  destroy(): void {
    this.el.removeEventListener('beforeinput', this.onBeforeInput)
    this.el.removeEventListener('compositionstart', this.onCompStart)
    this.el.removeEventListener('compositionupdate', this.onCompUpdate)
    this.el.removeEventListener('compositionend', this.onCompEnd)
    this.el.removeEventListener('input', this.onInput)
    this.el.removeEventListener('keydown', this.onKeyDown)
    this.el.remove()
  }

  private onKeyDown = (e: KeyboardEvent): void => {
    this.handlers.onKeyDown?.(e)
  }

  private onCompStart = (): void => {
    this.composing = true
  }

  private onCompUpdate = (e: CompositionEvent): void => {
    this.handlers.onCompositionUpdate(e.data)
  }

  private onCompEnd = (e: CompositionEvent): void => {
    this.composing = false
    this.handlers.onCompositionEnd(e.data)
    this.el.value = ''
  }

  private onBeforeInput = (e: InputEvent): void => {
    if (this.composing) return
    if (e.inputType === 'insertText' && e.data) {
      e.preventDefault()
      this.handlers.onCommit(e.data)
      this.el.value = ''
    } else if (e.inputType === 'insertParagraph' || e.inputType === 'insertLineBreak') {
      e.preventDefault()
      this.handlers.onCommit('\n')
      this.el.value = ''
    }
  }

  private onInput = (): void => {
    if (this.composing) return
    if (this.el.value) {
      this.handlers.onCommit(this.el.value)
      this.el.value = ''
    }
  }
}
