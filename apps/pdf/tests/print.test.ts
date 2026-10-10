import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import {
  printPdf,
  printScaleForAreas,
  subsetPdfPages,
  PRINT_DIALOG_TIMEOUT_MS,
} from '@genoffice/pdf-print'

/** Mirror of PDF_VIEWER_READY_MS in print.ts (not exported). */
const PDF_VIEWER_READY_MS = 350

async function tinyPdf(pageCount: number): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (let i = 0; i < pageCount; i += 1) doc.addPage([612, 792])
  return doc.save()
}

function stubPrintIframe(printImpl: () => void): void {
  const originalAppend = document.body.appendChild.bind(document.body)
  vi.spyOn(document.body, 'appendChild').mockImplementation((node) => {
    const el = originalAppend(node)
    if (el instanceof HTMLIFrameElement) {
      const fakeWin = {
        print: vi.fn(() => printImpl()),
        addEventListener: (type: string, listener: EventListener) => {
          window.addEventListener(type, listener)
        },
        removeEventListener: (type: string, listener: EventListener) => {
          window.removeEventListener(type, listener)
        },
        dispatchEvent: (event: Event) => window.dispatchEvent(event),
      }
      Object.defineProperty(el, 'contentWindow', {
        configurable: true,
        get: () => fakeWin,
      })
      queueMicrotask(() => el.onload?.(new Event('load')))
    }
    return el
  })
}

beforeEach(() => {
  vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:print-fake')
  vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {})
})

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
  vi.useRealTimers()
})

describe('printPdf', () => {
  it('loads a PDF blob in a hidden iframe and opens the print dialog', async () => {
    const data = await tinyPdf(2)
    let printed = false
    stubPrintIframe(() => {
      printed = true
      window.dispatchEvent(new Event('afterprint'))
    })

    vi.useFakeTimers()
    const done = printPdf(data)
    await vi.advanceTimersByTimeAsync(PDF_VIEWER_READY_MS)
    await done

    expect(printed).toBe(true)
    expect(document.querySelector('iframe[title="print"]')).toBeNull()
    expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:print-fake')
  })

  it('refuses to open the print dialog when every requested page is out of range', async () => {
    const data = await tinyPdf(3)
    await expect(printPdf(data, [0])).rejects.toThrow('none of the requested pages')
    await expect(printPdf(data, [999])).rejects.toThrow('none of the requested pages')
    expect(document.querySelector('iframe[title="print"]')).toBeNull()
  })

  it('gives up when afterprint never fires', async () => {
    vi.useFakeTimers()
    const data = await tinyPdf(1)
    stubPrintIframe(() => {
      /* never fires afterprint */
    })

    const done = printPdf(data)
    const settled = done.then(
      () => 'resolved',
      (err: Error) => err.message,
    )
    await vi.advanceTimersByTimeAsync(PDF_VIEWER_READY_MS)
    await vi.advanceTimersByTimeAsync(PRINT_DIALOG_TIMEOUT_MS)
    await expect(settled).resolves.toContain('never reported completion')
    expect(document.querySelector('iframe[title="print"]')).toBeNull()
  })
})

describe('subsetPdfPages', () => {
  it('keeps only the requested pages in order', async () => {
    const data = await tinyPdf(5)
    const out = await subsetPdfPages(data, [4, 2, 2, 99])
    const doc = await PDFDocument.load(out)
    expect(doc.getPageCount()).toBe(2)
  })

  it('returns the original bytes when every page is selected', async () => {
    const data = await tinyPdf(2)
    const out = await subsetPdfPages(data, [1, 2])
    expect(out).toBe(data)
  })
})

/** US-Letter page area at scale 1 (PDF points²) */
const LETTER_AREA = 612 * 792

describe('printScaleForAreas', () => {
  it('uses the 150 DPI target while the document fits the pixel budget', () => {
    expect(printScaleForAreas([LETTER_AREA])).toBeCloseTo(150 / 72, 10)
    expect(printScaleForAreas(new Array(70).fill(LETTER_AREA))).toBeCloseTo(150 / 72, 10)
  })

  it('falls back to the target for empty or degenerate input', () => {
    expect(printScaleForAreas([])).toBeCloseTo(150 / 72, 10)
    expect(printScaleForAreas([0])).toBeCloseTo(150 / 72, 10)
  })

  it('scales down proportionally over budget, never below the 120 DPI floor', () => {
    const mid = printScaleForAreas(new Array(100).fill(LETTER_AREA))
    expect(mid).toBeLessThan(150 / 72)
    expect(mid).toBeGreaterThanOrEqual(120 / 72)
    expect(printScaleForAreas(new Array(500).fill(LETTER_AREA))).toBeCloseTo(120 / 72, 10)
  })
})
