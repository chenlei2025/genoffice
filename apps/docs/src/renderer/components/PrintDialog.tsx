/**
 * Word print dialog: PDF page preview on the left, range on the right.
 * Builds the same PDF used for export, then prints via the PDF-app path.
 */
import { useCallback, useEffect, useRef, useState } from 'react'
import { GlobalWorkerOptions, getDocument, type PDFDocumentProxy } from 'pdfjs-dist/legacy/build/pdf.mjs'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'
import { t as tModule, useI18n } from '../i18n/locale'
import { parsePrintRange } from '../print-range'
import { clearPrintZoom, setPrintZoom } from '../print-zoom'

type RangeMode = 'all' | 'current' | 'custom'

let workerReady = false
function ensureWorker(): void {
  if (workerReady) return
  GlobalWorkerOptions.workerSrc = workerUrl
  workerReady = true
}

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return new Uint8Array(bytes)
}

const pvPages = (): HTMLElement[] => [
  ...document.querySelectorAll<HTMLElement>('.pagination-preview .pv-page'),
]

async function paintPdfPage(
  doc: PDFDocumentProxy,
  pageIndex0: number,
  canvas: HTMLCanvasElement,
  pane: HTMLElement,
): Promise<void> {
  const pdfPage = await doc.getPage(pageIndex0 + 1)
  try {
    const base = pdfPage.getViewport({ scale: 1 })
    const fit = Math.min(
      Math.max(pane.clientWidth - 32, 1) / base.width,
      Math.max(pane.clientHeight - 32, 1) / base.height,
    )
    const cssScale = Math.max(0.15, Math.min(3, fit))
    const dpr = Math.min(window.devicePixelRatio || 1, 2)
    const viewport = pdfPage.getViewport({ scale: cssScale })
    canvas.width = Math.floor(viewport.width * dpr)
    canvas.height = Math.floor(viewport.height * dpr)
    canvas.style.width = `${Math.floor(viewport.width)}px`
    canvas.style.height = `${Math.floor(viewport.height)}px`
    const ctx = canvas.getContext('2d')
    if (ctx) {
      ctx.save()
      ctx.setTransform(1, 0, 0, 1, 0, 0)
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, canvas.width, canvas.height)
      ctx.restore()
    }
    await pdfPage.render({
      canvas,
      viewport,
      background: 'rgb(255,255,255)',
      transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
    }).promise
  } finally {
    try {
      pdfPage.cleanup()
    } catch {
      /* older pdf.js */
    }
  }
}

/** Lay out pagination-preview pages as one PDF (same chunk/merge path as export). */
async function buildPreviewPdf(
  els: HTMLElement[],
  scale: number,
): Promise<{ ok: true; base64: string } | { ok: false; error: string }> {
  const pxToTwips = (px: number) => Math.round((px / 96) * 1440)
  type Group = { w: number; h: number; from: number; to: number }
  let groups: Group[] = []
  els.forEach((page, i) => {
    const w = pxToTwips(parseFloat(page.style.width))
    const h = pxToTwips(parseFloat(page.style.height))
    const last = groups[groups.length - 1]
    if (last && last.w === w && last.h === h) last.to = i
    else groups.push({ w, h, from: i, to: i })
  })
  const PRINT_CHUNK = 10
  if (els.length > PRINT_CHUNK) {
    const chunked: Group[] = []
    for (const g of groups) {
      for (let s = g.from; s <= g.to; s += PRINT_CHUNK) {
        chunked.push({ w: g.w, h: g.h, from: s, to: Math.min(s + PRINT_CHUNK - 1, g.to) })
      }
    }
    groups = chunked
  }
  const parts: string[] = []
  for (const g of groups) {
    els.forEach((page, i) => {
      page.classList.toggle('pv-print-skip', i < g.from || i > g.to)
      page.classList.toggle('pv-print-tail', i === g.to)
    })
    const part = await window.desktop.printPdfBuffer(g.w, g.h, scale)
    if (!part.ok || !part.base64) {
      return { ok: false, error: part.error ?? 'pdf' }
    }
    parts.push(part.base64)
  }
  const merged =
    parts.length === 1
      ? { ok: true as const, base64: parts[0]! }
      : await window.desktop.mergePdfBase64(parts)
  if (!merged.ok || !merged.base64) {
    return { ok: false, error: merged.error ?? 'merge' }
  }
  return { ok: true, base64: merged.base64 }
}

export function PrintDialog({
  onClose,
  setStatus,
}: {
  onClose: () => void
  setStatus: (s: string) => void
}) {
  const { t } = useI18n()
  const [printers, setPrinters] = useState<Array<{ name: string; displayName: string }>>([])
  const [deviceName, setDeviceName] = useState('')
  const [pageCount, setPageCount] = useState(0)
  const [page, setPage] = useState(0)
  const [rangeMode, setRangeMode] = useState<RangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [printing, setPrinting] = useState(false)
  const [loading, setLoading] = useState(true)
  const docRef = useRef<PDFDocumentProxy | null>(null)
  const base64Ref = useRef<string | null>(null)
  const taskDestroyRef = useRef<(() => void) | null>(null)
  const paneRef = useRef<HTMLDivElement | null>(null)
  const canvasRef = useRef<HTMLCanvasElement | null>(null)
  const onCloseRef = useRef(onClose)
  const setStatusRef = useRef(setStatus)
  onCloseRef.current = onClose
  setStatusRef.current = setStatus

  // Wait for pagination pages, then build a PDF and open it for preview/print.
  useEffect(() => {
    let cancelled = false
    let last = -1
    let stableTicks = 0
    const timer = window.setInterval(() => {
      const n = pvPages().length
      if (n > 0 && n === last) stableTicks += 1
      else stableTicks = 0
      last = n
      if (n === 0 || stableTicks < 2) return
      window.clearInterval(timer)
      void (async () => {
        const els = pvPages()
        if (els.length === 0 || cancelled) return
        const scale = setPrintZoom()
        try {
          const built = await buildPreviewPdf(els, scale)
          if (cancelled) return
          if (!built.ok) {
            setStatusRef.current(tModule('appPrintFailed', { error: built.error }))
            onCloseRef.current()
            return
          }
          ensureWorker()
          const task = getDocument({ data: decodeBase64(built.base64), useWasm: false })
          taskDestroyRef.current = () => {
            void task.destroy()
          }
          const doc = await task.promise
          if (cancelled) {
            void task.destroy()
            return
          }
          docRef.current = doc
          base64Ref.current = built.base64
          setPageCount(doc.numPages)
          setLoading(false)
        } catch (err) {
          if (!cancelled) {
            setStatusRef.current(
              tModule('appPrintFailed', {
                error: err instanceof Error ? err.message : String(err),
              }),
            )
            onCloseRef.current()
          }
        } finally {
          clearPrintZoom()
          els.forEach((el) => el.classList.remove('pv-print-skip', 'pv-print-tail'))
        }
      })()
    }, 250)
    return () => {
      cancelled = true
      window.clearInterval(timer)
      docRef.current = null
      base64Ref.current = null
      taskDestroyRef.current?.()
      taskDestroyRef.current = null
      clearPrintZoom()
      pvPages().forEach((el) => el.classList.remove('pv-print-skip', 'pv-print-tail'))
    }
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onCloseRef.current()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [])

  useEffect(() => {
    let cancelled = false
    void window.desktop.listPrinters().then((listed) => {
      if (cancelled || !Array.isArray(listed)) return
      setPrinters(listed)
      setDeviceName((prev) => prev || listed[0]?.name || '')
    })
    return () => {
      cancelled = true
    }
  }, [])

  const current = Math.min(page, Math.max(0, pageCount - 1))
  const ready = pageCount > 0 && !loading

  const selected: number[] | null =
    rangeMode === 'all'
      ? Array.from({ length: pageCount }, (_x, i) => i)
      : rangeMode === 'current'
        ? ready
          ? [current]
          : []
        : parsePrintRange(customRange, pageCount)
  const rangeInvalid = rangeMode === 'custom' && selected === null

  const renderPage = useCallback(async () => {
    const doc = docRef.current
    const canvas = canvasRef.current
    const pane = paneRef.current
    if (!doc || !canvas || !pane || pageCount === 0) return
    if (pane.clientWidth < 8 || pane.clientHeight < 8) return
    try {
      await paintPdfPage(doc, current, canvas, pane)
    } catch {
      /* cancelled */
    }
  }, [current, pageCount])

  useEffect(() => {
    if (!ready) return
    void renderPage()
  }, [ready, renderPage])

  useEffect(() => {
    const pane = paneRef.current
    if (!pane || !ready) return
    const ro = new ResizeObserver(() => {
      void renderPage()
    })
    ro.observe(pane)
    return () => ro.disconnect()
  }, [ready, renderPage])

  const doPrint = async () => {
    const pdfBase64 = base64Ref.current
    if (!selected || selected.length === 0 || printing || !pdfBase64) return
    setPrinting(true)
    try {
      const result = await window.desktop.printPdfSilent(
        pdfBase64,
        selected.map((index) => index + 1),
        deviceName || undefined,
      )
      if (!result.ok) {
        if (result.error) {
          setStatus(t('appPrintFailed', { error: result.error }))
        }
        return
      }
      onClose()
    } catch (err) {
      setStatus(t('appPrintFailed', { error: err instanceof Error ? err.message : String(err) }))
    } finally {
      setPrinting(false)
    }
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal print-dialog" onClick={(e) => e.stopPropagation()}>
        <h2>{t('appPrintTitle')}</h2>
        <div className="print-dialog-body">
          <div className="print-preview-pane">
            {ready ? (
              <>
                <div className="print-preview-page" ref={paneRef}>
                  <canvas ref={canvasRef} className="print-preview-canvas" />
                </div>
                <div className="print-preview-nav">
                  <button
                    type="button"
                    aria-label={t('appPrintPrevPage')}
                    title={t('appPrintPrevPage')}
                    disabled={current <= 0}
                    onClick={() => setPage(current - 1)}
                  >
                    ‹
                  </button>
                  <span>
                    {current + 1} / {pageCount}
                  </span>
                  <button
                    type="button"
                    aria-label={t('appPrintNextPage')}
                    title={t('appPrintNextPage')}
                    disabled={current >= pageCount - 1}
                    onClick={() => setPage(current + 1)}
                  >
                    ›
                  </button>
                </div>
              </>
            ) : (
              <div className="print-preview-empty">{t('appPrintRendering')}</div>
            )}
          </div>
          <div className="print-options">
            {printers.length > 0 && (
              <label className="print-printer">
                <span>{t('appPrintPrinter')}</span>
                <select
                  value={deviceName}
                  onChange={(e) => setDeviceName(e.target.value)}
                  disabled={printing}
                >
                  {printers.map((p) => (
                    <option key={p.name} value={p.name}>
                      {p.displayName}
                    </option>
                  ))}
                </select>
              </label>
            )}
            <fieldset>
              <legend>{t('appPrintRange')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-range"
                  checked={rangeMode === 'all'}
                  onChange={() => setRangeMode('all')}
                />
                {t('appPrintRangeAll')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-range"
                  checked={rangeMode === 'current'}
                  onChange={() => setRangeMode('current')}
                />
                {t('appPrintRangeCurrent')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-range"
                  checked={rangeMode === 'custom'}
                  onChange={() => setRangeMode('custom')}
                />
                {t('appPrintRangeCustom')}
              </label>
              <input
                type="text"
                className={`print-range-input ${rangeInvalid ? 'invalid' : ''}`}
                placeholder={t('appPrintRangeHint')}
                value={customRange}
                disabled={rangeMode !== 'custom'}
                onFocus={() => setRangeMode('custom')}
                onChange={(e) => setCustomRange(e.target.value)}
              />
            </fieldset>
            <div className="print-page-count">
              {ready && selected && selected.length > 0
                ? t('appPrintPageCount', { n: selected.length })
                : ''}
            </div>
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button
            type="button"
            className="primary"
            disabled={!ready || !selected || selected.length === 0 || printing || rangeInvalid}
            onClick={() => void doPrint()}
          >
            {printing ? t('appPrintProgress') : t('appPrintTitle')}
          </button>
        </div>
      </div>
    </div>
  )
}
