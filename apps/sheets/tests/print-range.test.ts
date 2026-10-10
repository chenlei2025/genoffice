import { describe, expect, it } from 'vitest'
import { parsePrintRange } from '../src/renderer/print-range'

describe('parsePrintRange', () => {
  it('parses commas and ranges', () => {
    expect(parsePrintRange('1,3,5-8', 10)).toEqual([0, 2, 4, 5, 6, 7])
  })

  it('dedupes and sorts', () => {
    expect(parsePrintRange('5-6, 2, 5', 10)).toEqual([1, 4, 5])
  })

  it('accepts CJK separators and en-dashes', () => {
    expect(parsePrintRange('1，3、5–6', 6)).toEqual([0, 2, 4, 5])
  })

  it('rejects empty or out-of-bounds input', () => {
    expect(parsePrintRange('', 5)).toBeNull()
    expect(parsePrintRange('0', 5)).toBeNull()
    expect(parsePrintRange('6', 5)).toBeNull()
    expect(parsePrintRange('4-2', 5)).toBeNull()
  })
})
