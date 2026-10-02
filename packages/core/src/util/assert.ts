export class AlmadocxError extends Error {
  readonly code: string

  constructor(code: string, message: string) {
    super(message)
    this.name = 'AlmadocxError'
    this.code = code
  }
}

export function assert(condition: unknown, code: string, message: string): asserts condition {
  if (!condition) {
    throw new AlmadocxError(code, message)
  }
}

export function unreachable(value: never, message = 'unreachable'): never {
  throw new AlmadocxError('unreachable', `${message}: ${String(value)}`)
}
