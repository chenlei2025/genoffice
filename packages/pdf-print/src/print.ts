import { PDFDocument } from 'pdf-lib'

/** How long to wait for the dialog to report back before giving up. */
export const PRINT_DIALOG_TIMEOUT_MS = 300_000

/** Chromium's PDF plugin needs a short beat after iframe load before print works. */
const PDF_VIEWER_READY_MS = 350

/**
 * Legacy helper kept for callers/tests that budget raster jobs. Native PDF
 * print no longer rasterizes; the scale is unused by {@link printPdf}.
 * @deprecated Prefer printing PDF bytes directly.
 */
export function printScaleForAreas(areas: number[]): number {
  const target = 150 / 72
  const total = areas.reduce((sum, area) => sum + area, 0)
  if (!(total > 0) || !Number.isFinite(total)) return target
  const budget = 150_000_000
  const atTarget = total * target * target
  if (!Number.isFinite(atTarget) || atTarget <= 0 || atTarget <= budget) return target
  return Math.max(120 / 72, target * Math.sqrt(budget / atTarget))
}

/** Copy selected 1-based pages into a new PDF (document order, de-duped). */
export async function subsetPdfPages(
  data: Uint8Array,
  pages: readonly number[],
): Promise<Uint8Array> {
  const src = await PDFDocument.load(data)
  const count = src.getPageCount()
  const indices = [
    ...new Set(
      pages.filter((n) => Number.isInteger(n) && n >= 1 && n <= count).map((n) => n - 1),
    ),
  ].sort((a, b) => a - b)
  if (indices.length === 0) {
    throw new Error(
      `none of the requested pages (${pages.join(', ')}) exist in this ${count}-page document`,
    )
  }
  if (indices.length === count) return data
  const out = await PDFDocument.create()
  const copied = await out.copyPages(src, indices)
  for (const page of copied) out.addPage(page)
  return out.save()
}

function waitForPrintDialog(timeoutMs: number, print: () => void, target: Window): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | null = null
    const detach = () => {
      if (timer !== null) clearTimeout(timer)
      timer = null
      target.removeEventListener('afterprint', onAfterPrint)
      window.removeEventListener('afterprint', onAfterPrint)
    }
    function onAfterPrint(): void {
      detach()
      resolve()
    }
    timer = setTimeout(() => {
      detach()
      reject(new Error('the print dialog never reported completion; released the PDF viewer'))
    }, timeoutMs)
    target.addEventListener('afterprint', onAfterPrint)
    // Some Electron builds bubble afterprint only on the embedding window.
    if (target !== window) window.addEventListener('afterprint', onAfterPrint)
    try {
      print()
    } catch (err) {
      detach()
      reject(err instanceof Error ? err : new Error(String(err)))
    }
  })
}

/**
 * Open the system print dialog on PDF bytes via Chromium's native PDF viewer.
 * Avoids rasterizing pages to JPEG (which made the dialog preview scroll jank).
 * @param pages 1-based file pages; defaults to the whole document.
 */
export async function printPdf(data: Uint8Array, pages?: number[]): Promise<void> {
  const bytes =
    pages && pages.length > 0 ? await subsetPdfPages(data, pages) : new Uint8Array(data)
  // Copy into a fresh ArrayBuffer — pdf-lib views may share a pool, and Blob
  // typing rejects SharedArrayBuffer-backed slices.
  const ab = new ArrayBuffer(bytes.byteLength)
  new Uint8Array(ab).set(bytes)
  const url = URL.createObjectURL(new Blob([ab], { type: 'application/pdf' }))
  const iframe = document.createElement('iframe')
  iframe.title = 'print'
  // Non-zero size: a 0×0 frame never finishes loading Chromium's PDF plugin.
  iframe.style.cssText =
    'position:fixed;left:-10000px;top:0;width:820px;height:1100px;border:0;opacity:0;pointer-events:none'
  document.body.appendChild(iframe)
  try {
    await new Promise<void>((resolve, reject) => {
      const timer = window.setTimeout(
        () => reject(new Error('the PDF viewer never finished loading')),
        60_000,
      )
      iframe.onload = () => {
        window.clearTimeout(timer)
        resolve()
      }
      iframe.onerror = () => {
        window.clearTimeout(timer)
        reject(new Error('the PDF viewer failed to load'))
      }
      iframe.src = url
    })
    await new Promise<void>((resolve) => {
      window.setTimeout(resolve, PDF_VIEWER_READY_MS)
    })
    const win = iframe.contentWindow
    if (!win) throw new Error('the PDF viewer has no content window')
    await waitForPrintDialog(PRINT_DIALOG_TIMEOUT_MS, () => win.print(), win)
  } finally {
    iframe.remove()
    URL.revokeObjectURL(url)
  }
}
