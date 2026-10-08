/// Screen page card and on-page alignment for the print preview. Print media
/// keeps the alignment and drops the gray preview chrome, so a pinch zoom
/// applied only under `@media screen` never changes the printed page.

export const PREVIEW_PAGE_SIZES = ['A3', 'A4', 'A5', 'Legal', 'Letter', 'Tabloid'] as const

export type PreviewPageSize =
  | (typeof PREVIEW_PAGE_SIZES)[number]
  | { width: number; height: number }

export type PreviewFitMode = 'scale' | 'sheet' | 'width'

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
}

/// Numbers the preview writes onto the page as CSS variables. Inline updates
/// on `<html>` override the baked stylesheet, so margin, scale, and alignment
/// edits repaint without rebuilding the style tag.
export function previewLayoutPatch(layout: PreviewLayout): PreviewLayoutPatch {
  const paper = typeof layout.pageSize === 'string' ? PAPER_INCHES[layout.pageSize] : layout.pageSize
  const across = layout.landscape ? paper.height : paper.width
  const down = layout.landscape ? paper.width : paper.height
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
  /* macOS print-dialog mat: dark chrome around a white page card. */
  html, body { margin: 0; height: 100%; background: #2c2c2e; color: #f5f5f7; }
  #preview-fit {
    height: 100%;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    align-items: stretch;
    box-sizing: border-box;
  }
  #preview-scroll {
    flex: 1;
    min-height: 0;
    overflow: auto;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 16px 16px 8px;
    box-sizing: border-box;
  }
  #preview-pages { display: flex; flex-direction: column; gap: 16px; align-items: center; }
  #preview-page-label {
    flex: none;
    text-align: center;
    padding: 8px 16px 12px;
    font: 12px/1.2 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    opacity: 0.85;
    user-select: none;
  }
  .area {
    position: relative;
    box-sizing: border-box;
    background: #fff;
    color: #000;
    width: calc(var(--page-w) * 1in);
    height: calc(var(--page-h) * 1in);
    padding: calc(var(--m-top) * 1in) calc(var(--m-right) * 1in) calc(var(--m-bottom) * 1in) calc(var(--m-left) * 1in);
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.45);
    display: flex;
    flex: none;
    flex-direction: column;
    justify-content: var(--align-y);
    align-items: var(--align-x);
    overflow: hidden;
  }
  .page-clip {
    flex: 1 1 auto;
    min-height: 0;
    width: 100%;
    overflow: hidden;
    position: relative;
  }
  /* Base print scale; __fitSheets further shrinks so the sheet fits the page. */
  .sheet { zoom: var(--print-scale); }
}
@media print {
  html, body, #preview-fit, #preview-scroll, #preview-pages {
    background: transparent !important;
    color: inherit !important;
    height: auto !important;
    display: block !important;
    overflow: visible !important;
    zoom: 1 !important;
  }
  #preview-page-label { display: none !important; }
  .page-clip {
    display: block !important;
    overflow: visible !important;
    height: auto !important;
  }
  .sheet { zoom: 1 !important; position: static !important; top: auto !important; }
  .area {
    background: transparent !important;
    width: auto !important;
    height: auto !important;
    min-height: var(--v-min) !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
    overflow: visible !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: var(--align-y) !important;
    align-items: var(--align-x) !important;
  }
}`
}
