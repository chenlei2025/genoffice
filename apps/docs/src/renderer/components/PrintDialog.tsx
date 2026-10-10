/**
 * Word print: build the same PDF as export (with pv-exporting so printToPDF
 * does not stall), then open the Sheets combined print dialog in main.
 */
import { useEffect, useRef } from 'react'
import { t as tModule, useI18n } from '../i18n/locale'
import { clearPrintZoom, setPrintZoom } from '../print-zoom'

const pvPages = (): HTMLElement[] => [
  ...document.querySelectorAll<HTMLElement>('.pagination-preview .pv-page'),
]

type Group = { w: number; h: number; from: number; to: number }

function chunkGroups(groups: Group[], size: number): Group[] {
  const chunked: Group[] = []
  for (const g of groups) {
    for (let s = g.from; s <= g.to; s += size) {
      chunked.push({ w: g.w, h: g.h, from: s, to: Math.min(s + size - 1, g.to) })
    }
  }
  return chunked
}

/** Same chunk/merge path as exportPdf — must use `.pv-exporting` or Chromium can stall for minutes. */
async function buildPreviewPdf(
  els: HTMLElement[],
  scale: number,
): Promise<{ ok: true; base64: string } | { ok: false; error: string }> {
  const pxToTwips = (px: number) => Math.round((px / 96) * 1440)
  let groups: Group[] = []
  els.forEach((page, i) => {
    const w = pxToTwips(parseFloat(page.style.width))
    const h = pxToTwips(parseFloat(page.style.height))
    const last = groups[groups.length - 1]
    if (last && last.w === w && last.h === h) last.to = i
    else groups.push({ w, h, from: i, to: i })
  })
  const PRINT_CHUNK = 10
  if (els.length > PRINT_CHUNK) groups = chunkGroups(groups, PRINT_CHUNK)

  const parts: string[] = []
  const appRoot = document.querySelector('.app')
  appRoot?.classList.add('pv-exporting')
  try {
    const queue = [...groups]
    while (queue.length > 0) {
      const g = queue.shift() as Group
      els.forEach((page, i) => {
        page.classList.toggle('pv-print-skip', i < g.from || i > g.to)
        page.classList.toggle('pv-print-tail', i === g.to)
      })
      const part = await window.desktop.printPdfBuffer(g.w, g.h, scale)
      if (part.ok && part.base64) {
        parts.push(part.base64)
        continue
      }
      if (g.to > g.from) {
        const mid = Math.floor((g.from + g.to) / 2)
        queue.unshift({ ...g, to: mid }, { ...g, from: mid + 1 })
        continue
      }
      return { ok: false, error: part.error ?? 'pdf' }
    }
  } finally {
    appRoot?.classList.remove('pv-exporting')
    els.forEach((el) => el.classList.remove('pv-print-skip', 'pv-print-tail'))
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
  fileName,
}: {
  onClose: () => void
  setStatus: (s: string) => void
  fileName: string
}) {
  const { t } = useI18n()
  const onCloseRef = useRef(onClose)
  const setStatusRef = useRef(setStatus)
  onCloseRef.current = onClose
  setStatusRef.current = setStatus

  useEffect(() => {
    let cancelled = false
    let last = -1
    let stableTicks = 0
    let waitedTicks = 0
    const timer = window.setInterval(() => {
      waitedTicks += 1
      const n = pvPages().length
      if (n > 0 && n === last) stableTicks += 1
      else stableTicks = 0
      last = n
      if (n === 0 || stableTicks < 2) {
        if (waitedTicks >= 80) {
          window.clearInterval(timer)
          if (!cancelled) {
            setStatusRef.current(
              tModule('appPrintFailed', { error: tModule('appPrintNoPages') }),
            )
            onCloseRef.current()
          }
        }
        return
      }
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
          clearPrintZoom()
          /// Dismiss this loader before the native dialog; unmount sets
          /// `cancelled`, so errors from printCombined are handled below.
          onCloseRef.current()
          try {
            const result = await window.desktop.printCombined(built.base64, fileName)
            if (!result.ok && result.error) {
              setStatusRef.current(tModule('appPrintFailed', { error: result.error }))
            }
          } catch (err) {
            setStatusRef.current(
              tModule('appPrintFailed', {
                error: err instanceof Error ? err.message : String(err),
              }),
            )
          }
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
          document.querySelector('.app')?.classList.remove('pv-exporting')
          els.forEach((el) => el.classList.remove('pv-print-skip', 'pv-print-tail'))
        }
      })()
    }, 250)
    return () => {
      cancelled = true
      window.clearInterval(timer)
      clearPrintZoom()
      document.querySelector('.app')?.classList.remove('pv-exporting')
      pvPages().forEach((el) => el.classList.remove('pv-print-skip', 'pv-print-tail'))
    }
  }, [fileName])

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

  /// Do not use `print-dialog` here: that class hides `.pagination-preview` for
  /// screen media, which is the printToPDF source.
  return (
    <div className="modal-backdrop print-loading-backdrop" onClick={onClose}>
      <div className="modal print-loading" onClick={(e) => e.stopPropagation()}>
        <p className="print-preview-empty">{t('appPrintRendering')}</p>
      </div>
    </div>
  )
}
