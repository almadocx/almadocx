import { describe, expect, it, vi } from 'vitest'
import { InputProxy } from '../src/input/ime.js'

describe('InputProxy', () => {
  it('commits insertText / line breaks and ignores beforeinput while composing', () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const onCommit = vi.fn()
    const onCompositionUpdate = vi.fn()
    const onCompositionEnd = vi.fn()
    const onKeyDown = vi.fn()
    const proxy = new InputProxy(host, {
      onCommit,
      onCompositionUpdate,
      onCompositionEnd,
      onKeyDown,
    })

    proxy.focus()
    expect(document.activeElement).toBe(proxy.el)
    expect(proxy.isComposing()).toBe(false)

    const insert = new InputEvent('beforeinput', {
      bubbles: true,
      cancelable: true,
      inputType: 'insertText',
      data: 'a',
    })
    proxy.el.dispatchEvent(insert)
    expect(onCommit).toHaveBeenCalledWith('a')

    onCommit.mockClear()
    proxy.el.dispatchEvent(
      new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertParagraph',
      }),
    )
    expect(onCommit).toHaveBeenCalledWith('\n')

    onCommit.mockClear()
    proxy.el.dispatchEvent(
      new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertLineBreak',
      }),
    )
    expect(onCommit).toHaveBeenCalledWith('\n')

    onCommit.mockClear()
    proxy.el.dispatchEvent(
      new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertText',
        data: null,
      }),
    )
    expect(onCommit).not.toHaveBeenCalled()

    proxy.el.dispatchEvent(new CompositionEvent('compositionstart'))
    expect(proxy.isComposing()).toBe(true)
    proxy.el.dispatchEvent(
      new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertText',
        data: 'ignored',
      }),
    )
    expect(onCommit).not.toHaveBeenCalled()

    proxy.el.dispatchEvent(new CompositionEvent('compositionupdate', { data: 'ni' }))
    expect(onCompositionUpdate).toHaveBeenCalledWith('ni')

    proxy.el.value = 'leftover'
    proxy.el.dispatchEvent(new CompositionEvent('compositionend', { data: '你' }))
    expect(proxy.isComposing()).toBe(false)
    expect(onCompositionEnd).toHaveBeenCalledWith('你')
    expect(proxy.el.value).toBe('')

    proxy.el.dispatchEvent(new KeyboardEvent('keydown', { key: 'a' }))
    expect(onKeyDown).toHaveBeenCalled()

    proxy.destroy()
    expect(host.contains(proxy.el)).toBe(false)
    host.remove()
  })

  it('falls back to input event commit when not composing', () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const onCommit = vi.fn()
    const proxy = new InputProxy(host, {
      onCommit,
      onCompositionUpdate: () => {},
      onCompositionEnd: () => {},
    })

    proxy.el.value = 'typed'
    proxy.el.dispatchEvent(new Event('input'))
    expect(onCommit).toHaveBeenCalledWith('typed')
    expect(proxy.el.value).toBe('')

    onCommit.mockClear()
    proxy.el.dispatchEvent(new CompositionEvent('compositionstart'))
    proxy.el.value = 'comp'
    proxy.el.dispatchEvent(new Event('input'))
    expect(onCommit).not.toHaveBeenCalled()

    proxy.el.value = ''
    proxy.el.dispatchEvent(new CompositionEvent('compositionend', { data: '' }))
    proxy.el.dispatchEvent(new Event('input'))
    expect(onCommit).not.toHaveBeenCalled()

    proxy.destroy()
    host.remove()
  })
})
