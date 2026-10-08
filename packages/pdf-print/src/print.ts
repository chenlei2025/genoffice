import type { PDFDocumentProxy } from 'pdfjs-dist'

/** Baseline raster density (previous behavior — quality never drops below this). */
const BASE_DPI = 150
/** Target raster density for sharper glyph edges on high-DPI printers. */
const TARGET_DPI = 200
/**
 * Upper bound on total raster pixels across all pages. A 200 DPI US-Letter
 * page is ~3.7 MP, so ~40 such pages print at full target quality; larger
 * documents scale down toward (never below) the 150 DPI baseline.
 */
const MAX_PRINT_PIXELS = 150_000_000

/** How long to wait for the dialog to report back before giving up. */
export const PRINT_DIALOG_TIMEOUT_MS = 300_000

const PRINT_STYLE_ID = 'genoffice-pdf-print-styles'

/**
 * Render scale for a document given each page's area at scale 1 (PDF points²).
 */
export function printScaleForAreas(areas: number[]): number {
  const target = TARGET_DPI / 72
  const total = areas.reduce((sum, area) => sum + area, 0)
  if (!(total > 0) || !Number.isFinite(total)) return target
  const atTarget = total * target * target
  if (!Number.isFinite(atTarget) || atTarget <= 0) return target
  if (atTarget <= MAX_PRINT_PIXELS) return target
  return Math.max(BASE_DPI / 72, target * Math.sqrt(MAX_PRINT_PIXELS / atTarget))
}

function ensurePrintStyles(): void {
  if (document.getElementById(PRINT_STYLE_ID)) return
  const style = document.createElement('style')
  style.id = PRINT_STYLE_ID
  style.textContent = `
.pdf-print-root { display: none; }
@media print {
  html, body { height: auto !important; overflow: visible !important; }
  body > *:not(.pdf-print-root) { display: none !important; }
  .pdf-print-root { display: block !important; }
  .pdf-print-root img {
    display: block;
    width: 100%;
    max-height: 100vh;
    object-fit: contain;
  }
  .pdf-print-root img + img { break-before: page; }
}
`
  document.head.appendChild(style)
}

function waitForPrintDialog(timeoutMs: number): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | null = null
    const detach = () => {
      if (timer !== null) clearTimeout(timer)
      timer = null
      window.removeEventListener('afterprint', onAfterPrint)
    }
    function onAfterPrint(): void {
      detach()
      resolve()
    }
    timer = setTimeout(() => {
      detach()
      reject(new Error('the print dialog never reported completion; released the rendered pages'))
    }, timeoutMs)
    window.addEventListener('afterprint', onAfterPrint)
    try {
      window.print()
    } catch (err) {
      detach()
      reject(err instanceof Error ? err : new Error(String(err)))
    }
  })
}

/**
 * Sequentially render pages as JPEG images into a print-only container, then
 * open the system print dialog (PDF app print path).
 * @param pages 1-based file pages; defaults to the whole document.
 */
export async function printPdf(doc: PDFDocumentProxy, pages?: number[]): Promise<void> {
  ensurePrintStyles()
  const root = document.createElement('div')
  root.className = 'pdf-print-root'
  const canvas = document.createElement('canvas')
  const targets =
    pages && pages.length > 0
      ? [...new Set(pages)]
          .filter((n) => Number.isInteger(n) && n >= 1 && n <= doc.numPages)
          .sort((a, b) => a - b)
      : Array.from({ length: doc.numPages }, (_x, i) => i + 1)
  if (pages && pages.length > 0 && targets.length === 0) {
    throw new Error(
      `none of the requested pages (${pages.join(', ')}) exist in this ${doc.numPages}-page document`,
    )
  }
  const areas: number[] = []
  for (const n of targets) {
    const page = await doc.getPage(n)
    try {
      const unit = page.getViewport({ scale: 1 })
      const area = unit.width * unit.height
      areas.push(Number.isFinite(area) && area > 0 ? area : 0)
    } finally {
      try {
        page.cleanup()
      } catch {
        /* older pdf.js without cleanup */
      }
    }
  }
  const scale = printScaleForAreas(areas)
  const MAX_PAGE_PIXELS = MAX_PRINT_PIXELS
  for (const n of targets) {
    const page = await doc.getPage(n)
    try {
      let viewport = page.getViewport({ scale })
      let w = Math.floor(viewport.width)
      let h = Math.floor(viewport.height)
      if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) continue
      if (w * h > MAX_PAGE_PIXELS) {
        const pageScale = scale * Math.sqrt(MAX_PAGE_PIXELS / (w * h))
        viewport = page.getViewport({ scale: pageScale })
        w = Math.floor(viewport.width)
        h = Math.floor(viewport.height)
        if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) continue
      }
      canvas.width = w
      canvas.height = h
      await page.render({ canvas, viewport }).promise
      const img = document.createElement('img')
      img.src = canvas.toDataURL('image/jpeg', 0.92)
      root.appendChild(img)
    } finally {
      try {
        page.cleanup()
      } catch {
        /* older pdf.js without cleanup */
      }
    }
  }
  canvas.width = 0
  canvas.height = 0
  document.body.appendChild(root)
  try {
    await Promise.all([...root.querySelectorAll('img')].map((img) => img.decode()))
    await waitForPrintDialog(PRINT_DIALOG_TIMEOUT_MS)
  } finally {
    root.remove()
  }
}

