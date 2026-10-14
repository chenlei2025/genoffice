/**
 * Per-file reading position persistence (WPS-style "return to where you left
 * off"). All entries live under one localStorage key as a path-keyed map with
 * LRU pruning, so the store cannot grow without bound. Pure module — storage
 * is injectable for tests.
 */

export type PdfFitMode = 'width' | 'page' | null

export type PdfViewState = {
  /** 1-based page number of the row at the viewport top */
  page: number
  /** Viewport-top offset within that row, as a fraction of the row height.
      Fraction (not pixels) so the position survives a different window size
      or fit-computed scale on the next open. */
  frac: number
  scale: number
  fitMode: PdfFitMode
}

type StoredEntry = PdfViewState & { at: number }

export const VIEW_STATE_KEY = 'genoffice-pdf-view-state'
export const MAX_VIEW_ENTRIES = 100

type StorageLike = Pick<Storage, 'getItem' | 'setItem'>

function readAll(storage: StorageLike): Record<string, unknown> {
  try {
    const parsed: unknown = JSON.parse(storage.getItem(VIEW_STATE_KEY) ?? '{}')
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
      ? (parsed as Record<string, unknown>)
      : {}
  } catch {
    return {}
  }
}

/** Entries come from disk and may be from an older build or corrupted — never trust them. */
function sanitize(entry: unknown): StoredEntry | null {
  if (!entry || typeof entry !== 'object') return null
  const e = entry as Record<string, unknown>
  if (typeof e.page !== 'number' || !Number.isFinite(e.page) || e.page < 1) return null
  if (typeof e.scale !== 'number' || !Number.isFinite(e.scale) || e.scale <= 0) return null
  if (e.fitMode !== 'width' && e.fitMode !== 'page' && e.fitMode !== null) return null
  return {
    page: Math.floor(e.page),
    frac:
      typeof e.frac === 'number' && Number.isFinite(e.frac) ? Math.min(Math.max(e.frac, 0), 1) : 0,
    scale: e.scale,
    fitMode: e.fitMode,
    at: typeof e.at === 'number' && Number.isFinite(e.at) ? e.at : 0,
  }
}

export function loadViewState(
  path: string,
  storage: StorageLike = localStorage,
): PdfViewState | null {
  if (!path) return null
  const entry = sanitize(readAll(storage)[path])
  if (!entry) return null
  const { at: _at, ...state } = entry
  return state
}

export function saveViewState(
  path: string,
  state: PdfViewState,
  storage: StorageLike = localStorage,
  now: number = Date.now(),
): void {
  if (!path) return
  const all = readAll(storage)
  all[path] = { ...state, at: now } satisfies StoredEntry
  const paths = Object.keys(all)
  if (paths.length > MAX_VIEW_ENTRIES) {
    const age = (p: string) => sanitize(all[p])?.at ?? -1
    paths.sort((a, b) => age(b) - age(a))
    for (const stale of paths.slice(MAX_VIEW_ENTRIES)) delete all[stale]
  }
  try {
    storage.setItem(VIEW_STATE_KEY, JSON.stringify(all))
  } catch {
    // Quota exceeded / storage unavailable: restoring is best-effort, never break viewing
  }
}

/** Derive the persistable view state from the current scroll geometry.
 *  Mirrors the row layout in App.tsx: rows stack vertically with a leading
 *  and inter-row gap; `rowHeights` are the on-screen row heights in px. */
export function captureViewState(args: {
  scrollTop: number
  rowHeights: number[]
  /** 1-based page number of each row's first page */
  rowPages: number[]
  gap: number
  scale: number
  fitMode: PdfFitMode
}): PdfViewState {
  const { scrollTop, rowHeights, rowPages, gap, scale, fitMode } = args
  let top = gap
  let rowIdx = 0
  let rowTop = top
  for (let i = 0; i < rowHeights.length; i++) {
    if (top <= scrollTop) {
      rowIdx = i
      rowTop = top
    } else break
    top += rowHeights[i]! + gap
  }
  const height = rowHeights[rowIdx] ?? 0
  return {
    page: rowPages[rowIdx] ?? 1,
    frac: height > 0 ? Math.min(Math.max((scrollTop - rowTop) / height, 0), 1) : 0,
    scale,
    fitMode,
  }
}

/** Row under a document point, page-content offset inside it at scale 1, and the fixed
 *  margin/gap remainder that never scales. `rowHeights` are at scale 1. */
export type ZoomAnchor = { rowIdx: number; content: number; fixed: number }

export function captureZoomAnchor(args: {
  y: number
  rowHeights: number[]
  gap: number
  scale: number
}): ZoomAnchor | null {
  const { y, rowHeights, gap, scale } = args
  if (rowHeights.length === 0 || scale <= 0) return null
  let top = gap
  let rowIdx = 0
  let rowTop = top
  for (let i = 0; i < rowHeights.length; i++) {
    if (top <= y) {
      rowIdx = i
      rowTop = top
    } else break
    top += rowHeights[i]! * scale + gap
  }
  const within = y - rowTop
  const content = Math.min(Math.max(within, 0), rowHeights[rowIdx]! * scale)
  return { rowIdx, content: content / scale, fixed: within - content }
}

/** Document-space y of a captured anchor once the rows are laid out at `scale`. */
export function zoomAnchorY(
  anchor: ZoomAnchor,
  rowHeights: number[],
  gap: number,
  scale: number,
): number {
  let top = gap
  for (let i = 0; i < anchor.rowIdx && i < rowHeights.length; i++)
    top += rowHeights[i]! * scale + gap
  return top + anchor.content * scale + anchor.fixed
}

/** Mount only rows near the viewport for long documents (spacers keep scroll height). */
export function rowWindowForScroll(args: {
  scrollTop: number
  clientHeight: number
  rowHeights: number[]
  gap: number
  scale: number
  buffer: number
}): { from: number; to: number; topPad: number; bottomPad: number } {
  const { scrollTop, clientHeight, rowHeights, gap, scale, buffer } = args
  const n = rowHeights.length
  if (n === 0) return { from: 0, to: -1, topPad: 0, bottomPad: 0 }
  const viewBottom = scrollTop + clientHeight
  // Match .pdf-page CSS heights (Math.floor(size * scale)) so spacers don't drift
  const px = (i: number) => Math.floor(rowHeights[i]! * scale)
  let y = gap
  let from = 0
  let to = 0
  for (let i = 0; i < n; i++) {
    const h = px(i)
    if (y + h >= scrollTop) {
      from = i
      break
    }
    y += h + gap
    from = i
  }
  y = gap
  for (let i = 0; i < n; i++) {
    const h = px(i)
    if (y <= viewBottom) to = i
    else break
    y += h + gap
  }
  from = Math.max(0, from - buffer)
  to = Math.min(n - 1, to + buffer)
  let topPad = 0
  for (let i = 0; i < from; i++) topPad += gap + px(i)
  let bottomPad = 0
  for (let i = to + 1; i < n; i++) bottomPad += gap + px(i)
  // Final document margin-bottom (last .pdf-row is not :last-child when a spacer follows)
  if (to < n - 1) bottomPad += gap
  return { from, to, topPad, bottomPad }
}
