/// Screen page card and on-page alignment for the print preview. Print media
/// keeps the alignment and drops the gray preview chrome, so a pinch zoom
/// applied only under `@media screen` never changes the printed page.

export const PREVIEW_PAGE_SIZES = ['A3', 'A4', 'A5', 'Legal', 'Letter', 'Tabloid'] as const

export type PreviewPageSize =
  | (typeof PREVIEW_PAGE_SIZES)[number]
  | { width: number; height: number }

export type PreviewFitMode = 'scale' | 'sheet' | 'width'

/// N-up fill order on each physical sheet (Layout pane).
export type PreviewLayoutDirection = 'ltr-ttb' | 'rtl-ttb' | 'ttb-ltr' | 'ttb-rtl'

export interface PreviewLayout {
  landscape: boolean
  pageSize: PreviewPageSize
  margins: { top: number; bottom: number; left: number; right: number }
  /// 0.1–2 when fitMode is `scale`.
  scale: number
  /// `sheet` / `width` shrink into the printable area like WPS “调整为”.
  fitMode: PreviewFitMode
  horizontalCentered: boolean
  verticalCentered: boolean
  /// Layout pane toggles (live preview + print job).
  reverseOrientation?: boolean
  flipHorizontal?: boolean
  pageBorder?: 'none' | 'hairline'
  pagesPerSheet?: number
  layoutDirection?: PreviewLayoutDirection
}

/// Chromium-supported N-up grids (1 / 2 / 4 / 6 / 9 / 16). Other values snap
/// up to the next supported size so the preview matches typical print engines.
export function nupGridSize(pagesPerSheet: number): { cols: number; rows: number } {
  const n = Math.min(16, Math.max(1, Math.round(pagesPerSheet) || 1))
  const supported = [1, 2, 4, 6, 9, 16] as const
  const hit = supported.find((value) => value >= n) ?? 16
  switch (hit) {
    case 2:
      return { cols: 2, rows: 1 }
    case 4:
      return { cols: 2, rows: 2 }
    case 6:
      return { cols: 3, rows: 2 }
    case 9:
      return { cols: 3, rows: 3 }
    case 16:
      return { cols: 4, rows: 4 }
    default:
      return { cols: 1, rows: 1 }
  }
}

/// Maps fill order index → row-major cell index for the Layout direction control.
export function nupCellIndex(
  fillIndex: number,
  cols: number,
  rows: number,
  direction: PreviewLayoutDirection,
): number {
  const total = cols * rows
  if (fillIndex < 0 || fillIndex >= total || cols < 1 || rows < 1) return fillIndex
  let row = 0
  let col = 0
  switch (direction) {
    case 'rtl-ttb':
      row = Math.floor(fillIndex / cols)
      col = cols - 1 - (fillIndex % cols)
      break
    case 'ttb-ltr':
      col = Math.floor(fillIndex / rows)
      row = fillIndex % rows
      break
    case 'ttb-rtl':
      col = cols - 1 - Math.floor(fillIndex / rows)
      row = fillIndex % rows
      break
    default:
      row = Math.floor(fillIndex / cols)
      col = fillIndex % cols
      break
  }
  return row * cols + col
}

const PAPER_INCHES: Record<(typeof PREVIEW_PAGE_SIZES)[number], { width: number; height: number }> =
  {
    Letter: { width: 8.5, height: 11 },
    Tabloid: { width: 11, height: 17 },
    Legal: { width: 8.5, height: 14 },
    A3: { width: 11.69, height: 16.54 },
    A4: { width: 8.27, height: 11.69 },
    A5: { width: 5.83, height: 8.27 },
  }

function round3(value: number): number {
  return Math.round(value * 1000) / 1000
}

export interface PreviewLayoutPatch {
  pageW: number
  pageH: number
  top: number
  right: number
  bottom: number
  left: number
  scale: number
  fitMode: PreviewFitMode
  align: 'center' | 'flex-start'
  justify: 'center' | 'flex-start'
  reverseOrientation: boolean
  flipHorizontal: boolean
  pageBorder: 'none' | 'hairline'
  pagesPerSheet: number
  layoutDirection: PreviewLayoutDirection
  nupCols: number
  nupRows: number
}

/// Numbers the preview writes onto the page as CSS variables. Inline updates
/// on `<html>` override the baked stylesheet, so margin, scale, and alignment
/// edits repaint without rebuilding the style tag.
export function previewLayoutPatch(layout: PreviewLayout): PreviewLayoutPatch {
  const paper = typeof layout.pageSize === 'string' ? PAPER_INCHES[layout.pageSize] : layout.pageSize
  const across = layout.landscape ? paper.height : paper.width
  const down = layout.landscape ? paper.width : paper.height
  const direction: PreviewLayoutDirection =
    layout.layoutDirection === 'rtl-ttb' ||
    layout.layoutDirection === 'ttb-ltr' ||
    layout.layoutDirection === 'ttb-rtl'
      ? layout.layoutDirection
      : 'ltr-ttb'
  const grid = nupGridSize(layout.pagesPerSheet ?? 1)
  return {
    pageW: round3(across),
    pageH: round3(down),
    top: round3(layout.margins.top),
    right: round3(layout.margins.right),
    bottom: round3(layout.margins.bottom),
    left: round3(layout.margins.left),
    scale: round3(Math.min(2, Math.max(0.1, layout.scale))),
    fitMode: layout.fitMode,
    align: layout.horizontalCentered ? 'center' : 'flex-start',
    justify: layout.verticalCentered ? 'center' : 'flex-start',
    reverseOrientation: layout.reverseOrientation === true,
    flipHorizontal: layout.flipHorizontal === true,
    pageBorder: layout.pageBorder === 'hairline' ? 'hairline' : 'none',
    pagesPerSheet: grid.cols * grid.rows,
    layoutDirection: direction,
    nupCols: grid.cols,
    nupRows: grid.rows,
  }
}

/// CSS that draws each print area as a full sheet of paper. On screen the
/// dialog scales `#preview-pages` so that sheet fits the preview; print media
/// drops that frame so the job uses the real paper size and margins.
/// Padding, scale, and alignment read variables so a live edit can override them.
export function previewLayoutCss(layout: PreviewLayout): string {
  const patch = previewLayoutPatch(layout)
  // Print content box ≈ paper minus Chromium margins; match that for vertical
  // center instead of 100vh (which does not equal a printed page).
  const vMin =
    patch.justify === 'center'
      ? `calc((${patch.pageH} - ${patch.top} - ${patch.bottom}) * 1in)`
      : '0'
  return `html {
  --page-w: ${patch.pageW}; --page-h: ${patch.pageH};
  --m-top: ${patch.top}; --m-right: ${patch.right}; --m-bottom: ${patch.bottom}; --m-left: ${patch.left};
  --print-scale: ${patch.scale};
  --align-x: ${patch.align}; --align-y: ${patch.justify}; --v-min: ${vMin};
}
.sheet { position: relative; width: max-content; flex: none; }
@media screen {
  /* Match the print dialog sidebar (Canvas) so left/right share one mat. */
  html, body { margin: 0; height: 100%; background: Canvas; color: CanvasText; }
  #preview-fit {
    height: 100%;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    align-items: stretch;
    box-sizing: border-box;
    background: Canvas;
  }
  #preview-scroll {
    flex: 1;
    min-height: 0;
    overflow: auto;
    display: block;
    padding: 16px 14px 10px;
    box-sizing: border-box;
    -webkit-overflow-scrolling: touch;
    overscroll-behavior: contain;
    overflow-anchor: none;
    scrollbar-width: thin;
    scrollbar-color: color-mix(in srgb, CanvasText 26%, transparent) transparent;
    /* Paint-isolate the scroller (no size containment — that can collapse scrollHeight). */
    contain: layout paint style;
    transform: translateZ(0);
  }
  #preview-scroll::-webkit-scrollbar { width: 8px; }
  #preview-scroll::-webkit-scrollbar-thumb {
    background: color-mix(in srgb, CanvasText 26%, transparent);
    border-radius: 4px;
  }
  #preview-pages {
    display: flex;
    flex-direction: column;
    gap: 14px;
    align-items: center;
    /* No transform on the stack — a scaled ancestor disables content-visibility
       skipping and forces one giant layer while scrolling. */
  }
  .preview-sheet {
    flex: none;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 8px;
  }
  .preview-sheet.is-off .preview-card-slot {
    opacity: 0.42;
  }
  .preview-page-picks {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 6px;
  }
  .preview-page-pick {
    min-width: 28px;
    height: 22px;
    padding: 0 8px;
    border-radius: 999px;
    border: 1px solid color-mix(in srgb, CanvasText 22%, transparent);
    background: color-mix(in srgb, CanvasText 8%, Canvas);
    color: CanvasText;
    font: 11px/22px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    cursor: pointer;
  }
  .preview-page-pick[aria-pressed="true"] {
    background: AccentColor;
    border-color: AccentColor;
    color: Canvas;
  }
  .preview-card-slot {
    flex: none;
    position: relative;
    overflow: hidden;
    background: #fff;
    border-radius: 2px;
    box-shadow:
      0 0 0 1px color-mix(in srgb, CanvasText 14%, transparent),
      0 1px 2px color-mix(in srgb, CanvasText 10%, transparent),
      0 6px 18px color-mix(in srgb, CanvasText 12%, transparent);
    /* No content-visibility: it ghosted page-pick pills into neighboring cards. */
    contain: layout paint style;
  }
  .preview-card-slot[data-empty="1"] {
    /* Lightweight paper placeholder while the heavy sheet DOM is unmounted. */
    background: #fff;
  }
  .preview-card-bitmap {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    display: block;
    pointer-events: none;
    user-select: none;
  }
  #preview-page-label {
    flex: none;
    text-align: center;
    padding: 8px 14px 12px;
    font: 11px/1.3 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    letter-spacing: 0.02em;
    color: color-mix(in srgb, CanvasText 62%, transparent);
    background: Canvas;
    border-top: 1px solid color-mix(in srgb, CanvasText 10%, transparent);
    user-select: none;
  }
  .area {
    position: absolute;
    top: 0;
    left: 0;
    box-sizing: border-box;
    background: #fff;
    color: #000;
    width: calc(var(--page-w) * 1in);
    height: calc(var(--page-h) * 1in);
    padding: calc(var(--m-top) * 1in) calc(var(--m-right) * 1in) calc(var(--m-bottom) * 1in) calc(var(--m-left) * 1in);
    border: 0;
    display: flex;
    flex-direction: column;
    justify-content: var(--align-y);
    align-items: var(--align-x);
    overflow: hidden;
    transform-origin: top left;
    contain: layout paint;
  }
  .page-clip {
    flex: 1 1 auto;
    min-height: 0;
    width: 100%;
    overflow: hidden;
    position: relative;
  }
  /* Screen preview: transform (not zoom) so scroll stays on the compositor. */
  .sheet {
    transform-origin: top left;
    backface-visibility: hidden;
  }
}
/* Prefer this box via printToPDF preferCSSPageSize — matches the screen card. */
@page {
  size: ${patch.pageW}in ${patch.pageH}in;
  margin: 0;
}
@media print {
  /* White paper under cells — transparent pages rasterize as black in pdf.js. */
  html, body, #preview-fit, #preview-scroll, #preview-pages {
    background: #fff !important;
    color: inherit !important;
    height: auto !important;
    display: block !important;
    overflow: visible !important;
  }
  #preview-page-label, .preview-card-slot, .preview-sheet { display: contents !important; }
  .preview-page-picks { display: none !important; }
  .preview-card-bitmap { display: none !important; }
  .page-clip {
    display: block !important;
    overflow: hidden !important;
    height: auto !important;
    max-width: 100% !important;
  }
  .sheet {
    zoom: 1 !important;
    transform: none !important;
    position: static !important;
    top: auto !important;
    max-width: 100% !important;
  }
  .area {
    position: static !important;
    background: #fff !important;
    /* Fill Chromium's margin box (same width the preview card content uses).
       width:auto let wide sheets overflow and get clipped on both sides when
       centered; 100% + hidden matches the screen overflow:hidden card. */
    width: 100% !important;
    max-width: 100% !important;
    height: auto !important;
    min-height: var(--v-min) !important;
    padding: 0 !important;
    margin: 0 !important;
    border: 0 !important;
    box-shadow: none !important;
    overflow: hidden !important;
    box-sizing: border-box !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: var(--align-y) !important;
    align-items: var(--align-x) !important;
    transform: none !important;
  }
}`
}
