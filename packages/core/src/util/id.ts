let seq = 0

export function nextId(prefix: string): string {
  seq += 1
  return `${prefix}_${seq.toString(36)}`
}

export function resetIdsForTests(): void {
  seq = 0
}
