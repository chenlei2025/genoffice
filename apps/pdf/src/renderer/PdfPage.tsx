import { useCallback, useEffect, useRef, useState } from 'react'
import type { CSSProperties, RefObject } from 'react'
import { AnnotationMode, TextLayer } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'
import { pdfRectToCss, quadToRect } from './annotations'
import type { LocalMarkup, PageGeom } from './annotations'
import { MAX_PAGE_RENDER_PIXELS, PAGE_RENDER_DPR_CAP } from './view-config'

/** Which items in the container are within the (expanded) viewport — shared lazy-render basis
    for pages/thumbnails. Rebuild the observer when enabled flips (sidebar toggles unmount/remount the root) */
export function useVisibleSet(
  rootRef: RefObject<HTMLElement | null>,
  count: number,
  rootMargin: string,
  enabled = true,
): { visible: Set<number>; setItemRef: (idx: number) => (el: HTMLElement | null) => void } {
  const [visible, setVisible] = useState<Set<number>>(() => new Set())
  const itemRefs = useRef<(HTMLElement | null)[]>([])
  const ioRef = useRef<IntersectionObserver | null>(null)
  const pendingRef = useRef<Map<number, boolean>>(new Map())
  const rafRef = useRef(0)
  // Stable per-index callbacks: a new `(el) => …` every render makes React call
  // ref(null) then ref(el), which was clearing thumbnails on every parent update.
  const refCbCache = useRef(new Map<number, (el: HTMLElement | null) => void>())
  const scheduleFlush = useCallback(() => {
    if (rafRef.current) return
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = 0
      const pending = pendingRef.current
      if (pending.size === 0) return
      pendingRef.current = new Map()
      setVisible((prev) => {
        let changed = false
        const next = new Set(prev)
        for (const [idx, on] of pending) {
          if (on) {
            if (!next.has(idx)) {
              next.add(idx)
              changed = true
            }
          } else if (next.delete(idx)) {
            changed = true
          }
        }
        return changed ? next : prev
      })
    })
  }, [])

  useEffect(() => {
    const root = rootRef.current
    if (!enabled || !root || count === 0) {
      setVisible((prev) => (prev.size === 0 ? prev : new Set()))
      return
    }
    const io = new IntersectionObserver(
      (entries) => {
        for (const e of entries) {
          const idx = Number((e.target as HTMLElement).dataset.idx)
          if (!Number.isFinite(idx)) continue
          pendingRef.current.set(idx, e.isIntersecting)
        }
        scheduleFlush()
      },
      { root, rootMargin },
    )
    ioRef.current = io
    for (const el of itemRefs.current) if (el) io.observe(el)
    return () => {
      if (rafRef.current) cancelAnimationFrame(rafRef.current)
      rafRef.current = 0
      pendingRef.current.clear()
      io.disconnect()
      ioRef.current = null
    }
  }, [rootRef, count, rootMargin, enabled, scheduleFlush])

  useEffect(() => {
    // Drop stale callbacks when the list shrinks (page delete / reorder)
    for (const idx of refCbCache.current.keys()) {
      if (idx >= count) refCbCache.current.delete(idx)
    }
    if (itemRefs.current.length > count) itemRefs.current.length = count
  }, [count])

  const setItemRef = useCallback((idx: number) => {
    let cb = refCbCache.current.get(idx)
    if (!cb) {
      cb = (el: HTMLElement | null) => {
        const prev = itemRefs.current[idx]
        if (prev === el) return
        if (prev && ioRef.current) ioRef.current.unobserve(prev)
        itemRefs.current[idx] = el
        if (el && ioRef.current) ioRef.current.observe(el)
        else if (!el) {
          pendingRef.current.set(idx, false)
          scheduleFlush()
        }
      }
      refCbCache.current.set(idx, cb)
    }
    return cb
  }, [scheduleFlush])

  return { visible, setItemRef }
}

/** Single page: renders canvas + text layer (select/copy) when visible.
 *  Offscreen pages keep their last bitmap (windowing unmounts far rows) so scrolling
 *  back does not flash white while a new render is queued. */
export function PdfPage({
  doc,
  pageNo,
  scale,
  rotationDelta,
  visible,
  onRenderState,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  scale: number
  /** Unsaved rotation delta (clockwise degrees) */
  rotationDelta: number
  visible: boolean
  onRenderState: (doc: PDFDocumentProxy, pageNo: number, pending: boolean) => void
}) {
  const holderRef = useRef<HTMLDivElement>(null)
  const paintedRef = useRef<{
    doc: PDFDocumentProxy
    pageNo: number
    scale: number
    rotationDelta: number
  } | null>(null)
  useEffect(() => {
    const holder = holderRef.current
    if (!holder) return
    // Still needed for the post-save barrier even when we keep the bitmap.
    if (!visible) {
      onRenderState(doc, pageNo, false)
      return
    }
    const painted = paintedRef.current
    // Same doc/scale already on screen — do not tear down the canvas (avoids flicker
    // when IntersectionObserver briefly toggles while scrolling).
    if (
      painted &&
      painted.doc === doc &&
      painted.pageNo === pageNo &&
      painted.scale === scale &&
      painted.rotationDelta === rotationDelta &&
      holder.querySelector('canvas')
    ) {
      onRenderState(doc, pageNo, false)
      return
    }
    // Visibility may change while a save reload is running. Register newly visible
    // pages dynamically so global preview cleanup cannot outrun their bitmap swap.
    onRenderState(doc, pageNo, true)
    let cancelled = false
    let renderTask: RenderTask | null = null
    let idleId = 0
    let idleTimer = 0
    void (async () => {
      const page = await doc.getPage(pageNo)
      if (cancelled) return
      const viewport = page.getViewport({ scale, rotation: (page.rotate + rotationDelta) % 360 })
      // Cap DPR: hi-dpi full-page bitmaps dominate scroll cost on long docs.
      // Deep zoom trades dpr for the pixel budget instead — the page is already magnified.
      const dpr = Math.min(
        window.devicePixelRatio || 1,
        PAGE_RENDER_DPR_CAP,
        Math.sqrt(MAX_PAGE_RENDER_PIXELS / (viewport.width * viewport.height)),
      )
      // Paint offscreen, then swap — never clear the visible canvas mid-render.
      const canvas = document.createElement('canvas')
      canvas.width = Math.floor(viewport.width * dpr)
      canvas.height = Math.floor(viewport.height * dpr)
      canvas.style.width = `${Math.floor(viewport.width)}px`
      canvas.style.height = `${Math.floor(viewport.height)}px`
      renderTask = page.render({
        canvas,
        viewport,
        // FormLayer renders interactive widgets as HTML. Exclude their saved appearance
        // streams from the page bitmap or filled values are drawn twice after a reload.
        annotationMode: AnnotationMode.ENABLE_FORMS,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      })
      try {
        await renderTask.promise
      } catch {
        return // cancelled
      }
      if (cancelled) return
      // Bitmap first so scrolling stays responsive; text layer is deferred to idle.
      holder.replaceChildren(canvas)
      paintedRef.current = { doc, pageNo, scale, rotationDelta }
      onRenderState(doc, pageNo, false)
      const runTextLayer = () => {
        if (cancelled || !holder.contains(canvas)) return
        const textDiv = document.createElement('div')
        textDiv.className = 'textLayer'
        holder.appendChild(textDiv)
        const textLayer = new TextLayer({
          textContentSource: page.streamTextContent(),
          container: textDiv,
          viewport,
        })
        void textLayer.render().catch(() => {
          /* cancelled */
        })
      }
      const ric = (
        window as Window & {
          requestIdleCallback?: (cb: () => void, opts?: { timeout: number }) => number
        }
      ).requestIdleCallback
      if (typeof ric === 'function') idleId = ric(runTextLayer, { timeout: 500 })
      else idleTimer = window.setTimeout(runTextLayer, 64)
    })()
    return () => {
      cancelled = true
      renderTask?.cancel()
      if (idleId) {
        const cic = (
          window as Window & { cancelIdleCallback?: (id: number) => void }
        ).cancelIdleCallback
        cic?.(idleId)
      }
      if (idleTimer) window.clearTimeout(idleTimer)
    }
  }, [doc, pageNo, scale, rotationDelta, visible, onRenderState])
  return <div ref={holderRef} className="pdf-page-content" />
}

/** Overlay for unsaved markups. Purely visual (pointer-events: none) so the text
 *  underneath stays selectable; clicking is handled by the page-level hit test. */
export function MarkupOverlay({
  markups,
  geom,
  scale,
  selectedId,
}: {
  markups: LocalMarkup[]
  geom: PageGeom
  scale: number
  selectedId: string | null
}) {
  return (
    <>
      {markups.flatMap((m) =>
        m.quads.map((q, i) => {
          const [r, g, b] = m.color
          const style: CSSProperties = pdfRectToCss(geom, quadToRect(q), scale)
          if (m.type === 'highlight') {
            style.background = `rgba(${r * 255}, ${g * 255}, ${b * 255}, 0.4)`
          } else {
            const bar = `rgb(${r * 255}, ${g * 255}, ${b * 255})`
            if (m.type === 'underline') style.borderBottom = `2px solid ${bar}`
            else style.backgroundImage = `linear-gradient(${bar}, ${bar})`
          }
          return (
            <div
              key={`${m.id}-${i}`}
              className={`pdf-markup pdf-markup-${m.type}${m.id === selectedId ? ' pdf-markup-selected' : ''}`}
              style={style}
            />
          )
        }),
      )}
    </>
  )
}
