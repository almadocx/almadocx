import type { DocRange } from '../model/position.js'
import type { Document } from '../model/types.js'
import { applyOp } from './apply.js'
import type { AppliedOp, EditorOp } from './types.js'

export interface HistoryEntry {
  applied: AppliedOp
  /** Selection before this step was applied (restored on undo). */
  before: DocRange
  /** Selection after this step was applied (restored on redo). */
  after: DocRange
}

export interface HistoryState {
  undoStack: HistoryEntry[]
  redoStack: HistoryEntry[]
}

function cloneRange(range: DocRange): DocRange {
  return {
    anchor: { ...range.anchor, ...(range.anchor.cell ? { cell: { ...range.anchor.cell } } : {}) },
    focus: { ...range.focus, ...(range.focus.cell ? { cell: { ...range.focus.cell } } : {}) },
  }
}

export function createHistory(): HistoryState {
  return { undoStack: [], redoStack: [] }
}

function isBatch(applied: AppliedOp): applied is AppliedOp & { _batch: AppliedOp[] } {
  return Array.isArray((applied as AppliedOp & { _batch?: AppliedOp[] })._batch)
}

function makeBatch(steps: AppliedOp[]): AppliedOp {
  if (steps.length === 1) return steps[0]!
  return {
    forward: { type: 'insertText', position: { sectionIndex: 0, blockIndex: 0, offset: 0 }, text: '' },
    inverse: { type: 'insertText', position: { sectionIndex: 0, blockIndex: 0, offset: 0 }, text: '' },
    _batch: steps,
  } as AppliedOp & { _batch: AppliedOp[] }
}

export function dispatch(
  doc: Document,
  history: HistoryState,
  op: EditorOp,
  selection: { before: DocRange; after: DocRange } = {
    before: { anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 }, focus: { sectionIndex: 0, blockIndex: 0, offset: 0 } },
    after: { anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 }, focus: { sectionIndex: 0, blockIndex: 0, offset: 0 } },
  },
): { doc: Document; history: HistoryState } {
  const { doc: next, applied } = applyOp(doc, op)
  return {
    doc: next,
    history: {
      undoStack: [
        ...history.undoStack,
        {
          applied,
          before: cloneRange(selection.before),
          after: cloneRange(selection.after),
        },
      ],
      redoStack: [],
    },
  }
}

/** Apply many ops as one undo/redo step (Word/Docs paste, multi-para list, etc.). */
export function dispatchBatch(
  doc: Document,
  history: HistoryState,
  ops: EditorOp[],
  selection: { before: DocRange; after: DocRange } = {
    before: { anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 }, focus: { sectionIndex: 0, blockIndex: 0, offset: 0 } },
    after: { anchor: { sectionIndex: 0, blockIndex: 0, offset: 0 }, focus: { sectionIndex: 0, blockIndex: 0, offset: 0 } },
  },
): { doc: Document; history: HistoryState } {
  if (ops.length === 0) return { doc, history }
  let d = doc
  const steps: AppliedOp[] = []
  for (const op of ops) {
    const { doc: next, applied } = applyOp(d, op)
    d = next
    steps.push(applied)
  }
  return {
    doc: d,
    history: {
      undoStack: [
        ...history.undoStack,
        {
          applied: makeBatch(steps),
          before: cloneRange(selection.before),
          after: cloneRange(selection.after),
        },
      ],
      redoStack: [],
    },
  }
}

export function undo(
  doc: Document,
  history: HistoryState,
): { doc: Document; history: HistoryState; selection?: DocRange } {
  const entry = history.undoStack[history.undoStack.length - 1]
  if (!entry) return { doc, history }
  const applied = entry.applied
  let next = doc
  if (isBatch(applied)) {
    for (let i = applied._batch.length - 1; i >= 0; i--) {
      next = applyOp(next, applied._batch[i]!.inverse).doc
    }
  } else {
    next = applyOp(doc, applied.inverse).doc
  }
  return {
    doc: next,
    history: {
      undoStack: history.undoStack.slice(0, -1),
      redoStack: [...history.redoStack, entry],
    },
    selection: cloneRange(entry.before),
  }
}

export function redo(
  doc: Document,
  history: HistoryState,
): { doc: Document; history: HistoryState; selection?: DocRange } {
  const entry = history.redoStack[history.redoStack.length - 1]
  if (!entry) return { doc, history }
  const applied = entry.applied
  let next = doc
  let againApplied: AppliedOp = applied
  if (isBatch(applied)) {
    const redone: AppliedOp[] = []
    for (const step of applied._batch) {
      const r = applyOp(next, step.forward)
      next = r.doc
      redone.push(r.applied)
    }
    againApplied = makeBatch(redone)
  } else {
    const r = applyOp(doc, applied.forward)
    next = r.doc
    againApplied = r.applied
  }
  const again: HistoryEntry = {
    applied: againApplied,
    before: cloneRange(entry.before),
    after: cloneRange(entry.after),
  }
  return {
    doc: next,
    history: {
      undoStack: [...history.undoStack, again],
      redoStack: history.redoStack.slice(0, -1),
    },
    selection: cloneRange(entry.after),
  }
}

export function canUndo(history: HistoryState): boolean {
  return history.undoStack.length > 0
}

export function canRedo(history: HistoryState): boolean {
  return history.redoStack.length > 0
}
