/**
 * Parse a print range ("1,3,5-8") into sorted 0-based page indices.
 * Returns null on any invalid or out-of-bounds part.
 */
export function parsePrintRange(text: string, max: number): number[] | null {
  const out = new Set<number>()
  const tokenRe = /\d+\s*[-–—~～〜]\s*\d+|\d+/g
  const parts = text.match(tokenRe) ?? []
  if (parts.length === 0) return null
  if (text.replace(tokenRe, '').replace(/[,，、;；\s]/g, '') !== '') return null
  for (const part of parts) {
    const m = /^(\d+)\s*[-–—~～〜]\s*(\d+)$|^(\d+)$/.exec(part)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i - 1)
  }
  return [...out].sort((x, y) => x - y)
}

/// 1-based page numbers → "1-3,5,7-8".
export function compactPrintRange(pages: readonly number[]): string {
  const sorted = [...new Set(pages.filter((n) => Number.isInteger(n) && n >= 1))].sort(
    (a, b) => a - b,
  )
  if (sorted.length === 0) return ''
  const parts: string[] = []
  let start = sorted[0]!
  let prev = sorted[0]!
  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i]!
    if (cur === prev + 1) {
      prev = cur
      continue
    }
    parts.push(start === prev ? String(start) : `${start}-${prev}`)
    start = prev = cur
  }
  parts.push(start === prev ? String(start) : `${start}-${prev}`)
  return parts.join(',')
}
