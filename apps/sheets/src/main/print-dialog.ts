/// One print window: page preview on the left, printer and page setup on the
/// right. Printing goes straight to the chosen printer, so the system print
/// panel does not open as a second window.

import { randomUUID } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

import { BrowserWindow, WebContentsView, dialog, ipcMain } from 'electron'

import { getUiLang } from '@genoffice/i18n'
import { showSaveDialogWithMemory } from '@genoffice/electron-utils'

import { atomicWriteFile } from './atomic-write'
import { appStrings } from '../renderer/i18n/strings-app'
import { previewLayoutCss, previewLayoutPatch } from '../shared/preview-layout'
import {
  NAMED_PAGE_SIZES,
  SAVE_AS_PDF_DEVICE,
  isSaveAsPdfDevice,
  parsePreviewLayout,
  parsePrintDialogChoice,
  printOptionsFromDialog,
  printToPdfOptionsFromDialog,
} from './print-options'
import type { IpcMainEvent, WebContents } from 'electron'
import type { WorkbookPrintResult } from '../shared/desktop-api'
import type { PreviewLayout } from '../shared/preview-layout'
import type { PrintDialogChoice, PrintDialogDefaults } from './print-options'

const SIDEBAR_PX = 340

const PAPER_DIMS_MM: Record<(typeof NAMED_PAGE_SIZES)[number], { w: number; h: number }> = {
  A3: { w: 297, h: 420 },
  A4: { w: 210, h: 297 },
  A5: { w: 148, h: 210 },
  Letter: { w: 216, h: 279 },
  Legal: { w: 216, h: 356 },
  Tabloid: { w: 279, h: 432 },
}

interface ListedPrinter {
  name: string
  displayName: string
}

type DialogLabelKey =
  | 'appPrintDialogTitle'
  | 'appPrintPrinter'
  | 'appPrintSaveAsPdf'
  | 'appPrintPdf'
  | 'appPrintCopies'
  | 'appPrintNow'
  | 'appPrintNoPrinter'
  | 'appPrintFailed'
  | 'appExportPdfTitle'
  | 'appCancel'
  | 'appPrintProperties'
  | 'appPrintPages'
  | 'appPrintPageOf'
  | 'appPrintAllPages'
  | 'appPrintPageFrom'
  | 'appPrintPageTo'
  | 'appPrintScale'
  | 'appPrintScaleRatio'
  | 'appPrintAdjustTo'
  | 'appPrintFitSheetOnePage'
  | 'appPrintFitColumnsOnePage'
  | 'appPrintSheetSection'
  | 'appPrintContent'
  | 'appPrintActiveSheet'
  | 'appPrintNormalSize'
  | 'appPrintColorMode'
  | 'appPrintColor'
  | 'appPrintGrayscale'
  | 'appPrintSides'
  | 'appPrintOneSided'
  | 'appPrintLongEdge'
  | 'appPrintShortEdge'
  | 'appPrintCollate'
  | 'appPrintMarginTop'
  | 'appPrintMarginBottom'
  | 'appPrintMarginLeft'
  | 'appPrintMarginRight'
  | 'appPrintCustomPaper'
  | 'appPrintMm'
  | 'appOrientationLabel'
  | 'appPortrait'
  | 'appLandscape'
  | 'appSizeLabel'
  | 'appPrintMarginsInches'
  | 'appPrintAlign'
  | 'appPrintCenterHorizontal'
  | 'appPrintCenterVertical'
  | 'appPrintCenterHorizontalShort'
  | 'appPrintCenterVerticalShort'
  | 'appWidth'
  | 'appHeight'

function label(key: DialogLabelKey): string {
  const lang = getUiLang()
  const dict = appStrings[lang] ?? appStrings.en
  return dict[key]
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (ch) => {
    switch (ch) {
      case '&':
        return '&amp;'
      case '<':
        return '&lt;'
      case '>':
        return '&gt;'
      case '"':
        return '&quot;'
      default:
        return '&#39;'
    }
  })
}

function paperOptions(defaults: PrintDialogDefaults): string {
  const custom = typeof defaults.pageSize !== 'string'
  const unit = label('appPrintMm')
  const named = NAMED_PAGE_SIZES.map((name) => {
    const { w, h } = PAPER_DIMS_MM[name]
    const text = `${name} ${w} × ${h} ${unit}`
    return `<option value="${name}"${defaults.pageSize === name ? ' selected' : ''}>${escapeHtml(text)}</option>`
  }).join('')
  const customOption = `<option value="custom"${custom ? ' selected' : ''}>${escapeHtml(label('appPrintCustomPaper'))}</option>`
  return named + customOption
}

function panelHtml(printers: readonly ListedPrinter[], defaults: PrintDialogDefaults): string {
  const data = JSON.stringify({ printers, defaults }).replace(/</g, '\\u003c')
  const none = printers.length === 0
  const custom = typeof defaults.pageSize === 'object' ? defaults.pageSize : { width: 8.5, height: 11 }
  const scalePercent = Math.round(defaults.scale * 100)
  const inch = (value: number) => String(Math.round(value * 100) / 100)
  const fitMode = defaults.fitMode
  const scaleMode = fitMode === 'scale'
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
  html, body { margin: 0; height: 100%; background: Canvas; color: CanvasText; }
  [hidden] { display: none !important; }
  aside {
    position: fixed;
    top: 0;
    right: 0;
    bottom: 0;
    width: ${SIDEBAR_PX}px;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    background: Canvas;
    border-left: 1px solid color-mix(in srgb, CanvasText 18%, transparent);
    font: 13px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  }
  .fields {
    flex: 1;
    overflow: auto;
    display: flex;
    flex-direction: column;
    gap: 12px;
    padding: 14px 16px 16px;
  }
  h2 {
    margin: 8px 0 0;
    padding-top: 8px;
    border-top: 1px solid color-mix(in srgb, CanvasText 14%, transparent);
    font-size: 12px;
    font-weight: 600;
  }
  h2:first-child { border-top: 0; padding-top: 0; margin-top: 0; }
  label.field { display: flex; flex-direction: column; gap: 4px; }
  label.check { display: flex; flex-direction: row; align-items: center; gap: 8px; }
  select, input[type="number"], input[type="text"] {
    font: inherit;
    padding: 5px 8px;
    border-radius: 5px;
    border: 1px solid color-mix(in srgb, CanvasText 22%, transparent);
    background: Field;
    color: FieldText;
  }
  .row { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
  .row > span { flex: none; }
  .row > select, .row > input { flex: 1; min-width: 0; }
  .pair, .margins { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .hint { margin: 0; opacity: 0.75; font-size: 12px; }
  .pages { display: flex; flex-direction: column; gap: 6px; }
  .pages .check input[type="number"] { width: 3.2em; padding: 3px 4px; }
  .orient {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
  }
  .orient label {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 6px;
    padding: 8px 6px;
    border: 1px solid color-mix(in srgb, CanvasText 22%, transparent);
    border-radius: 8px;
    cursor: pointer;
  }
  .orient label:has(input:checked) {
    border-color: AccentColor;
    background: color-mix(in srgb, AccentColor 16%, transparent);
  }
  .orient input { position: absolute; opacity: 0; pointer-events: none; }
  .orient svg { display: block; }
  .center-row { display: flex; flex-wrap: wrap; gap: 12px 16px; }
  .scale-row { display: flex; flex-direction: column; gap: 8px; }
  .scale-row .check { flex-wrap: wrap; }
  .scale-row .check input[type="number"] { width: 4.2em; }
  .scale-row .check select { flex: 1; min-width: 0; }
  .toggle {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
  }
  .toggle input {
    appearance: none;
    width: 40px;
    height: 22px;
    border-radius: 11px;
    background: color-mix(in srgb, CanvasText 28%, transparent);
    position: relative;
    cursor: pointer;
    border: 0;
    flex: none;
  }
  .toggle input::after {
    content: '';
    position: absolute;
    top: 2px;
    left: 2px;
    width: 18px;
    height: 18px;
    border-radius: 50%;
    background: Canvas;
    transition: left 0.12s ease;
  }
  .toggle input:checked { background: AccentColor; }
  .toggle input:checked::after { left: 20px; }
  .actions {
    flex: none;
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    align-items: center;
    justify-content: space-between;
    padding: 12px 16px;
    border-top: 1px solid color-mix(in srgb, CanvasText 18%, transparent);
    background: Canvas;
  }
  .actions-end { display: flex; gap: 8px; margin-left: auto; }
  button { font: inherit; padding: 6px 14px; border-radius: 6px; }
  button.primary { background: AccentColor; color: AccentColorText; border: 0; }
  button:disabled { opacity: 0.5; }
</style>
</head>
<body>
<aside>
<div class="fields">
  <label class="field">${escapeHtml(label('appPrintPrinter'))}
    <select id="printer"${none ? ' disabled' : ''}></select>
  </label>
  <div class="row">
    <span>${escapeHtml(label('appPrintCopies'))}</span>
    <input id="copies" type="number" min="1" max="99" value="1" style="max-width:5em">
  </div>
  <div class="pages">
    <div>${escapeHtml(label('appPrintPages'))}</div>
    <label class="check">
      <input type="radio" name="pages" id="pages-all" value="all" checked>
      ${escapeHtml(label('appPrintAllPages'))}
    </label>
    <label class="check">
      <input type="radio" name="pages" id="pages-range" value="range">
      ${escapeHtml(label('appPrintPageFrom'))}
      <input id="page-from" type="number" min="1" max="9999" value="1" disabled>
      ${escapeHtml(label('appPrintPageTo'))}
      <input id="page-to" type="number" min="1" max="9999" value="1" disabled>
    </label>
  </div>
  <label class="toggle">
    <span>${escapeHtml(label('appPrintColorMode'))}</span>
    <input id="color" type="checkbox">
  </label>
  <label class="field">${escapeHtml(label('appSizeLabel'))}
    <select id="paper">${paperOptions(defaults)}</select>
  </label>
  <div class="pair" id="custom-paper"${typeof defaults.pageSize === 'string' ? ' hidden' : ''}>
    <label class="field">${escapeHtml(label('appWidth'))}
      <input id="paper-width" type="number" min="1" max="100" step="0.01" value="${custom.width}">
    </label>
    <label class="field">${escapeHtml(label('appHeight'))}
      <input id="paper-height" type="number" min="1" max="100" step="0.01" value="${custom.height}">
    </label>
  </div>
  <div>
    <div style="margin-bottom:6px">${escapeHtml(label('appOrientationLabel'))}</div>
    <div class="orient">
      <label>
        <input type="radio" name="orientation" id="orient-portrait" value="portrait"${defaults.landscape ? '' : ' checked'}>
        <svg width="28" height="36" viewBox="0 0 28 36" aria-hidden="true">
          <rect x="4" y="2" width="20" height="32" rx="2" fill="none" stroke="currentColor" stroke-width="1.5"/>
          <line x1="8" y1="10" x2="20" y2="10" stroke="currentColor" stroke-width="1"/>
          <line x1="8" y1="14" x2="20" y2="14" stroke="currentColor" stroke-width="1"/>
          <line x1="8" y1="18" x2="16" y2="18" stroke="currentColor" stroke-width="1"/>
        </svg>
        <span>${escapeHtml(label('appPortrait'))}</span>
      </label>
      <label>
        <input type="radio" name="orientation" id="orient-landscape" value="landscape"${defaults.landscape ? ' checked' : ''}>
        <svg width="36" height="28" viewBox="0 0 36 28" aria-hidden="true">
          <rect x="2" y="4" width="32" height="20" rx="2" fill="none" stroke="currentColor" stroke-width="1.5"/>
          <line x1="8" y1="10" x2="28" y2="10" stroke="currentColor" stroke-width="1"/>
          <line x1="8" y1="14" x2="28" y2="14" stroke="currentColor" stroke-width="1"/>
          <line x1="8" y1="18" x2="20" y2="18" stroke="currentColor" stroke-width="1"/>
        </svg>
        <span>${escapeHtml(label('appLandscape'))}</span>
      </label>
    </div>
  </div>

  <h2>${escapeHtml(label('appPrintSheetSection'))}</h2>
  <label class="field">${escapeHtml(label('appPrintContent'))}
    <select id="print-content" disabled>
      <option selected>${escapeHtml(label('appPrintActiveSheet'))}</option>
    </select>
  </label>
  <div>
    <div style="margin-bottom:6px">${escapeHtml(label('appPrintAlign'))}</div>
    <div class="center-row">
      <label class="check">
        <input id="center-h" type="checkbox"${defaults.horizontalCentered ? ' checked' : ''}>
        ${escapeHtml(label('appPrintCenterHorizontalShort'))}
      </label>
      <label class="check">
        <input id="center-v" type="checkbox"${defaults.verticalCentered ? ' checked' : ''}>
        ${escapeHtml(label('appPrintCenterVerticalShort'))}
      </label>
    </div>
  </div>
  <div class="scale-row">
    <label class="check">
      <input type="radio" name="scale-mode" id="scale-mode-pct" value="scale"${scaleMode ? ' checked' : ''}>
      ${escapeHtml(label('appPrintScaleRatio'))}
      <input id="scale" type="number" min="10" max="200" step="1" value="${scalePercent}"${scaleMode ? '' : ' disabled'}>
      <span>${escapeHtml(label('appPrintNormalSize'))}</span>
    </label>
    <label class="check">
      <input type="radio" name="scale-mode" id="scale-mode-fit" value="fit"${scaleMode ? '' : ' checked'}>
      ${escapeHtml(label('appPrintAdjustTo'))}
      <select id="fit-adjust"${scaleMode ? ' disabled' : ''}>
        <option value="sheet"${fitMode !== 'width' ? ' selected' : ''}>${escapeHtml(label('appPrintFitSheetOnePage'))}</option>
        <option value="width"${fitMode === 'width' ? ' selected' : ''}>${escapeHtml(label('appPrintFitColumnsOnePage'))}</option>
      </select>
    </label>
  </div>
  <div>
    <div style="margin-bottom:6px">${escapeHtml(label('appPrintMarginsInches'))}</div>
    <div class="margins">
      <label class="field">${escapeHtml(label('appPrintMarginTop'))}
        <input id="margin-top" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.top)}">
      </label>
      <label class="field">${escapeHtml(label('appPrintMarginBottom'))}
        <input id="margin-bottom" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.bottom)}">
      </label>
      <label class="field">${escapeHtml(label('appPrintMarginLeft'))}
        <input id="margin-left" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.left)}">
      </label>
      <label class="field">${escapeHtml(label('appPrintMarginRight'))}
        <input id="margin-right" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.right)}">
      </label>
    </div>
  </div>
  <label class="toggle">
    <span>${escapeHtml(label('appPrintCollate'))}</span>
    <input id="collate" type="checkbox" checked>
  </label>
  <label class="field">${escapeHtml(label('appPrintSides'))}
    <select id="duplex">
      <option value="simplex">${escapeHtml(label('appPrintOneSided'))}</option>
      <option value="longEdge">${escapeHtml(label('appPrintLongEdge'))}</option>
      <option value="shortEdge">${escapeHtml(label('appPrintShortEdge'))}</option>
    </select>
  </label>
  ${none ? `<p class="hint" id="no-printer">${escapeHtml(label('appPrintNoPrinter'))}</p>` : ''}
</div>
  <div class="actions">
    <button id="save-pdf" type="button" title="${escapeHtml(label('appPrintSaveAsPdf'))}">${escapeHtml(label('appPrintPdf'))}</button>
    <div class="actions-end">
      <button id="cancel" type="button">${escapeHtml(label('appCancel'))}</button>
      <button id="print" class="primary" type="button"${none ? ' disabled' : ''}>${escapeHtml(label('appPrintNow'))}</button>
    </div>
  </div>
</aside>
<script id="dialog-data" type="application/json">${data}</script>
<script>
  const data = JSON.parse(document.getElementById('dialog-data').textContent)
  const SAVE_AS_PDF = ${JSON.stringify(SAVE_AS_PDF_DEVICE)}
  const select = document.getElementById('printer')
  for (const printer of data.printers) {
    const option = document.createElement('option')
    option.value = printer.name
    option.textContent = printer.displayName || printer.name
    select.appendChild(option)
  }
  const pagesAll = document.getElementById('pages-all')
  const pagesRange = document.getElementById('pages-range')
  const pageFrom = document.getElementById('page-from')
  const pageTo = document.getElementById('page-to')
  const syncPages = () => {
    const rangeOn = pagesRange.checked
    pageFrom.disabled = !rangeOn
    pageTo.disabled = !rangeOn
  }
  pagesAll.onchange = syncPages
  pagesRange.onchange = syncPages
  syncPages()
  const paper = document.getElementById('paper')
  const customPaper = document.getElementById('custom-paper')
  paper.onchange = () => { customPaper.hidden = paper.value !== 'custom' }
  const scaleInput = document.getElementById('scale')
  const fitAdjust = document.getElementById('fit-adjust')
  const scaleModePct = document.getElementById('scale-mode-pct')
  const scaleModeFit = document.getElementById('scale-mode-fit')
  const syncScaleMode = () => {
    const pct = scaleModePct.checked
    scaleInput.disabled = !pct
    fitAdjust.disabled = pct
  }
  scaleModePct.onchange = syncScaleMode
  scaleModeFit.onchange = syncScaleMode
  syncScaleMode()
  const printBtn = document.getElementById('print')
  const num = (id) => Number(document.getElementById(id).value)
  const currentFitMode = () => {
    if (scaleModePct.checked) return 'scale'
    return fitAdjust.value === 'width' ? 'width' : 'sheet'
  }
  const landscape = () => document.getElementById('orient-landscape').checked
  const currentLayout = () => ({
    landscape: landscape(),
    pageSize: paper.value === 'custom'
      ? { width: num('paper-width'), height: num('paper-height') }
      : paper.value,
    margins: {
      top: num('margin-top'),
      bottom: num('margin-bottom'),
      left: num('margin-left'),
      right: num('margin-right'),
    },
    scale: (Math.round(num('scale')) || 100) / 100,
    fitMode: currentFitMode(),
    horizontalCentered: document.getElementById('center-h').checked,
    verticalCentered: document.getElementById('center-v').checked,
  })
  const publishLayout = () => window.printDialog.layout(currentLayout())
  for (const id of [
    'paper', 'orient-portrait', 'orient-landscape', 'scale', 'scale-mode-pct', 'scale-mode-fit',
    'fit-adjust', 'center-h', 'center-v', 'paper-width', 'paper-height',
    'margin-top', 'margin-bottom', 'margin-left', 'margin-right',
  ]) {
    const el = document.getElementById(id)
    el.addEventListener('input', publishLayout)
    el.addEventListener('change', publishLayout)
  }
  const savePdfBtn = document.getElementById('save-pdf')
  const submitChoice = (deviceName) => {
    if (!window.printDialog) return
    if (!deviceName) return
    printBtn.disabled = true
    savePdfBtn.disabled = true
    document.getElementById('cancel').disabled = true
    const pageSize = paper.value === 'custom'
      ? { width: num('paper-width'), height: num('paper-height') }
      : paper.value
    let pageRange = null
    if (pagesRange.checked) {
      let from = Math.round(num('page-from'))
      let to = Math.round(num('page-to'))
      if (from > to) { const swap = from; from = to; to = swap }
      pageRange = { from, to }
    }
    window.printDialog.submit({
      deviceName,
      copies: Math.round(num('copies')) || 1,
      collate: document.getElementById('collate').checked,
      color: document.getElementById('color').checked,
      duplexMode: document.getElementById('duplex').value,
      landscape: landscape(),
      pageSize,
      margins: {
        top: num('margin-top'),
        bottom: num('margin-bottom'),
        left: num('margin-left'),
        right: num('margin-right'),
      },
      scale: (Math.round(num('scale')) || 100) / 100,
      fitMode: currentFitMode(),
      pages: pageRange,
      horizontalCentered: document.getElementById('center-h').checked,
      verticalCentered: document.getElementById('center-v').checked,
    })
  }
  document.getElementById('cancel').onclick = () => {
    if (window.printDialog) window.printDialog.submit(null)
  }
  savePdfBtn.onclick = () => submitChoice(SAVE_AS_PDF)
  printBtn.onclick = () => submitChoice(select.value)
</script>
</body>
</html>`
}

function preloadSource(channel: string, layoutChannel: string): string {
  return `const { contextBridge, ipcRenderer } = require('electron')
contextBridge.exposeInMainWorld('printDialog', {
  submit: (choice) => ipcRenderer.send(${JSON.stringify(channel)}, choice),
  layout: (layout) => ipcRenderer.send(${JSON.stringify(layoutChannel)}, layout),
})
`
}

/// Builds stacked page cards for screen preview (vertical pagination), then
/// fits page width into the pane. Pinch zoom only changes the pane scale;
/// `__contentScale` is what print/PDF must use. Source HTML on disk is
/// unchanged — print still paginates via Chromium.
const PREVIEW_FIT_SCRIPT = `(() => {
  if (window.__fitPreview) return
  let userZoom = 1
  let baseScale = 1
  let fitMode = 'sheet'
  const pageOf = ${JSON.stringify(label('appPrintPageOf'))}
  const originals = [...document.querySelectorAll('.area')]
  const sources = originals.map((area) => area.cloneNode(true))
  for (const area of originals) area.remove()
  const fitHost = document.createElement('div')
  fitHost.id = 'preview-fit'
  const scroll = document.createElement('div')
  scroll.id = 'preview-scroll'
  const pages = document.createElement('div')
  pages.id = 'preview-pages'
  scroll.appendChild(pages)
  fitHost.appendChild(scroll)
  const pageLabel = document.createElement('div')
  pageLabel.id = 'preview-page-label'
  fitHost.appendChild(pageLabel)
  document.body.appendChild(fitHost)

  const applyPageChrome = (area) => {
    const root = document.documentElement
    const pageW = root.style.getPropertyValue('--page-w') || getComputedStyle(root).getPropertyValue('--page-w')
    const pageH = root.style.getPropertyValue('--page-h') || getComputedStyle(root).getPropertyValue('--page-h')
    const top = root.style.getPropertyValue('--m-top') || getComputedStyle(root).getPropertyValue('--m-top')
    const right = root.style.getPropertyValue('--m-right') || getComputedStyle(root).getPropertyValue('--m-right')
    const bottom = root.style.getPropertyValue('--m-bottom') || getComputedStyle(root).getPropertyValue('--m-bottom')
    const left = root.style.getPropertyValue('--m-left') || getComputedStyle(root).getPropertyValue('--m-left')
    const align = root.style.getPropertyValue('--align-x') || getComputedStyle(root).getPropertyValue('--align-x') || 'flex-start'
    const justify = root.style.getPropertyValue('--align-y') || getComputedStyle(root).getPropertyValue('--align-y') || 'flex-start'
    area.style.boxSizing = 'border-box'
    area.style.display = 'flex'
    area.style.flexDirection = 'column'
    if (pageW) area.style.width = pageW.trim() + 'in'
    if (pageH) area.style.height = pageH.trim() + 'in'
    area.style.padding = [top, right, bottom, left].map((v) => (v || '0').trim() + 'in').join(' ')
    area.style.justifyContent = justify.trim() || 'flex-start'
    area.style.alignItems = align.trim() || 'flex-start'
    area.style.overflow = 'hidden'
  }

  const updatePageLabel = () => {
    const cards = [...pages.querySelectorAll('.area')]
    const total = Math.max(1, cards.length)
    let current = 1
    if (cards.length > 0) {
      const mid = scroll.scrollTop + scroll.clientHeight * 0.35
      const zoom = parseFloat(pages.style.zoom) || 1
      for (let i = 0; i < cards.length; i++) {
        const top = cards[i].offsetTop * zoom
        const bottom = top + cards[i].offsetHeight * zoom
        if (mid >= top && mid < bottom) {
          current = i + 1
          break
        }
        if (mid >= bottom) current = i + 1
      }
    }
    pageLabel.textContent = pageOf
      .replace('{current}', String(current))
      .replace('{total}', String(total))
  }

  /// Rebuild stacked page cards from the source print areas.
  const rebuildPages = () => {
    pages.style.zoom = '1'
    pages.innerHTML = ''
    let contentScale = baseScale
    for (const source of sources) {
      const probe = source.cloneNode(true)
      pages.appendChild(probe)
      applyPageChrome(probe)
      const sheet = probe.querySelector('.sheet')
      if (!sheet) continue
      sheet.style.zoom = '1'
      sheet.style.position = ''
      sheet.style.top = ''
      sheet.style.flex = 'none'
      sheet.style.width = 'max-content'
      sheet.style.minWidth = 'max-content'
      const style = getComputedStyle(probe)
      const padX = (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0)
      const padY = (parseFloat(style.paddingTop) || 0) + (parseFloat(style.paddingBottom) || 0)
      const aw = Math.max(probe.clientWidth - padX, 1)
      const ah = Math.max(probe.clientHeight - padY, 1)
      const sw = Math.max(sheet.scrollWidth, sheet.offsetWidth, 1)
      const sh = Math.max(sheet.scrollHeight, sheet.offsetHeight, 1)
      let zoom = baseScale
      if (fitMode === 'width') zoom = Math.min(aw / sw, 1)
      else if (fitMode === 'sheet') zoom = Math.min(aw / sw, ah / sh, 1)
      zoom = Math.max(0.1, Math.min(2, zoom))
      contentScale = Math.min(contentScale, zoom)
      const scaledH = sh * zoom
      const pageCount = fitMode === 'sheet'
        ? 1
        : Math.max(1, Math.ceil(scaledH / ah - 1e-6))
      probe.remove()
      const alignX = (getComputedStyle(document.documentElement)
        .getPropertyValue('--align-x') || 'flex-start').trim() || 'flex-start'
      for (let index = 0; index < pageCount; index++) {
        const card = source.cloneNode(true)
        applyPageChrome(card)
        const cardSheet = card.querySelector('.sheet')
        if (!cardSheet) {
          pages.appendChild(card)
          continue
        }
        cardSheet.style.zoom = String(zoom)
        cardSheet.style.flex = 'none'
        cardSheet.style.width = 'max-content'
        cardSheet.style.minWidth = 'max-content'
        if (pageCount === 1) {
          pages.appendChild(card)
          continue
        }
        // Multi-page: clip each card to one printable slice of the sheet.
        card.style.justifyContent = 'flex-start'
        card.style.alignItems = alignX
        const clip = document.createElement('div')
        clip.className = 'page-clip'
        clip.style.display = 'flex'
        clip.style.justifyContent = alignX
        clip.style.alignItems = 'flex-start'
        cardSheet.style.position = 'relative'
        // top is pre-zoom CSS px; visual shift = top * zoom ≈ one printable page.
        cardSheet.style.top = index === 0 ? '0' : String(-index * ah / zoom) + 'px'
        cardSheet.replaceWith(clip)
        clip.appendChild(cardSheet)
        pages.appendChild(card)
      }
    }
    window.__contentScale = contentScale
    window.__fitMode = fitMode
  }

  const fit = () => {
    rebuildPages()
    const page = pages.querySelector('.area')
    if (!page) {
      updatePageLabel()
      return
    }
    const width = page.offsetWidth
    if (scroll.clientWidth < 40 || !width) {
      updatePageLabel()
      return
    }
    // Fit page width into the pane; height scrolls so every page stays readable.
    const fitted = (scroll.clientWidth - 32) / width
    pages.style.zoom = String(Math.max(0.2, Math.min(4, fitted * userZoom)))
    updatePageLabel()
  }
  window.__fitPreview = fit
  window.__setPrintScale = (scale, mode) => {
    baseScale = typeof scale === 'number' && scale > 0 ? scale : 1
    fitMode = mode === 'sheet' || mode === 'width' ? mode : 'scale'
    fit()
  }
  scroll.addEventListener('scroll', updatePageLabel, { passive: true })
  window.addEventListener('resize', fit)
  window.addEventListener('wheel', (event) => {
    if (!event.ctrlKey) return
    event.preventDefault()
    userZoom = Math.max(0.2, Math.min(5, userZoom * Math.exp(-event.deltaY * 0.01)))
    fit()
  }, { passive: false })
  fit()
})()`

function runPrint(
  contents: WebContents,
  options: ReturnType<typeof printOptionsFromDialog>,
): Promise<{ success: boolean; failureReason: string }> {
  return new Promise((resolve) => {
    contents.print(options, (success, failureReason) => resolve({ success, failureReason }))
  })
}

/// HTML for the real print/PDF job: the laid-out sheet plus the dialog's
/// current page setup. No preview fit wrappers or inline screen styles, so
/// the paper matches what the left pane simulates.
async function writeJobHtml(
  sourcePath: string,
  destPath: string,
  choice: Pick<
    PrintDialogChoice,
    | 'landscape'
    | 'pageSize'
    | 'margins'
    | 'scale'
    | 'fitMode'
    | 'horizontalCentered'
    | 'verticalCentered'
  >,
): Promise<void> {
  const layoutCss = previewLayoutCss({
    landscape: choice.landscape,
    pageSize: choice.pageSize,
    margins: choice.margins,
    scale: choice.scale,
    // Job window prints via scaleFactor; keep CSS on fixed scale.
    fitMode: 'scale',
    horizontalCentered: choice.horizontalCentered,
    verticalCentered: choice.verticalCentered,
  })
  const source = await readFile(sourcePath, 'utf8')
  const html = source.includes('</head>')
    ? source.replace('</head>', `<style id="go-preview-layout">${layoutCss}</style></head>`)
    : `${source}<style id="go-preview-layout">${layoutCss}</style>`
  await writeFile(destPath, html, 'utf8')
}

async function openJobWindow(htmlPath: string): Promise<BrowserWindow> {
  const hidden = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  await hidden.loadFile(htmlPath)
  return hidden
}

/// Render the dialog's page setup to a PDF file. Caller already picked `outPath`
/// (save panel must run before the modal print window is destroyed on macOS).
async function savePdfFromPrintDialog(args: {
  htmlPath: string
  workDir: string
  outPath: string
  choice: PrintDialogChoice
}): Promise<WorkbookPrintResult> {
  const pdfHtmlPath = join(args.workDir, 'print-save.pdf.html')
  await writeJobHtml(args.htmlPath, pdfHtmlPath, args.choice)
  const hidden = await openJobWindow(pdfHtmlPath)
  try {
    const pdf = await hidden.webContents.printToPDF(printToPdfOptionsFromDialog(args.choice))
    await atomicWriteFile(args.outPath, pdf)
    return { ok: true, path: args.outPath }
  } finally {
    if (!hidden.isDestroyed()) hidden.destroy()
  }
}

async function closePrintWindow(win: BrowserWindow): Promise<void> {
  if (win.isDestroyed()) return
  await new Promise<void>((resolve) => {
    win.once('closed', () => resolve())
    win.destroy()
  })
}

/// Silent print through a clean window so the job uses the same HTML/CSS as
/// Save as PDF — not the preview WebContentsView (fit zoom, gray mat, inline
/// screen padding), which made the paper diverge from the left pane.
async function printFromPrintDialog(args: {
  htmlPath: string
  workDir: string
  choice: PrintDialogChoice
}): Promise<WorkbookPrintResult> {
  const jobPath = join(args.workDir, 'print-job.html')
  await writeJobHtml(args.htmlPath, jobPath, args.choice)
  const hidden = await openJobWindow(jobPath)
  try {
    let outcome = await runPrint(hidden.webContents, printOptionsFromDialog(args.choice))
    if (!outcome.success && !/cancel/i.test(outcome.failureReason)) {
      outcome = await runPrint(hidden.webContents, printOptionsFromDialog(args.choice, true))
    }
    if (outcome.success) return { ok: true }
    return /cancel/i.test(outcome.failureReason)
      ? { ok: false }
      : { ok: false, error: outcome.failureReason }
  } finally {
    if (!hidden.isDestroyed()) hidden.destroy()
  }
}

export async function openCombinedPrintDialog(args: {
  parent: BrowserWindow | null
  htmlPath: string
  workDir: string
  settings: PrintDialogDefaults
  /// Suggested PDF file name when the destination is Save as PDF.
  fileName: string
}): Promise<WorkbookPrintResult> {
  const channel = `sheets:print-dialog:${randomUUID()}`
  const layoutChannel = `sheets:print-dialog-layout:${randomUUID()}`
  const preloadPath = join(args.workDir, 'print-dialog-preload.js')
  const panelPath = join(args.workDir, 'print-dialog.html')
  await writeFile(preloadPath, preloadSource(channel, layoutChannel), 'utf8')

  const parent = args.parent && !args.parent.isDestroyed() ? args.parent : undefined
  const win = new BrowserWindow({
    width: 1040,
    height: 760,
    minWidth: 820,
    minHeight: 520,
    show: false,
    title: label('appPrintDialogTitle'),
    autoHideMenuBar: true,
    ...(parent ? { parent, modal: true } : {}),
    webPreferences: { sandbox: true, preload: preloadPath },
  })
  const preview = new WebContentsView({
    webPreferences: { sandbox: true },
  })
  win.contentView.addChildView(preview)

  const refitPreview = (): void => {
    if (preview.webContents.isDestroyed()) return
    void preview.webContents.executeJavaScript('window.__fitPreview && window.__fitPreview()').catch(() => {})
  }
  const layout = (): void => {
    const [width = 0, height = 0] = win.getContentSize()
    preview.setBounds({
      x: 0,
      y: 0,
      width: Math.max(0, width - SIDEBAR_PX),
      height,
    })
    refitPreview()
  }
  win.on('resize', layout)

  let layoutGen = 0
  let fittedPage = ''
  /// Write margins, scale, and alignment onto the page itself. Inline styles
  /// beat the baked stylesheet so each keystroke moves the sheet inside the
  /// paper card without refitting (which would reset the pinch zoom).
  const applyPreviewLayout = (layout: PreviewLayout): Promise<void> => {
    if (preview.webContents.isDestroyed()) return Promise.resolve()
    const patch = previewLayoutPatch(layout)
    const pageKey = `${patch.pageW}x${patch.pageH}`
    const gen = ++layoutGen
    return preview.webContents
      .executeJavaScript(
        `(() => {
          const layout = ${JSON.stringify(patch)}
          const root = document.documentElement
          root.style.setProperty('--page-w', String(layout.pageW))
          root.style.setProperty('--page-h', String(layout.pageH))
          root.style.setProperty('--m-top', String(layout.top))
          root.style.setProperty('--m-right', String(layout.right))
          root.style.setProperty('--m-bottom', String(layout.bottom))
          root.style.setProperty('--m-left', String(layout.left))
          root.style.setProperty('--print-scale', String(layout.scale))
          root.style.setProperty('--align-x', layout.align)
          root.style.setProperty('--align-y', layout.justify)
          root.style.setProperty(
            '--v-min',
            layout.justify === 'center'
              ? 'calc((' + layout.pageH + ' - ' + layout.top + ' - ' + layout.bottom + ') * 1in)'
              : '0',
          )
          const pad = layout.top + 'in ' + layout.right + 'in ' + layout.bottom + 'in ' + layout.left + 'in'
          for (const area of document.querySelectorAll('.area')) {
            area.style.boxSizing = 'border-box'
            area.style.display = 'flex'
            area.style.flexDirection = 'column'
            area.style.width = layout.pageW + 'in'
            area.style.height = layout.pageH + 'in'
            area.style.padding = pad
            area.style.justifyContent = layout.justify
            area.style.alignItems = layout.align
            area.style.overflow = 'hidden'
          }
          for (const sheet of document.querySelectorAll('.sheet')) {
            sheet.style.flex = 'none'
            sheet.style.width = 'max-content'
            sheet.style.minWidth = 'max-content'
            sheet.style.maxWidth = 'none'
          }
          if (typeof window.__setPrintScale === 'function') {
            window.__setPrintScale(layout.scale, layout.fitMode)
          } else {
            for (const sheet of document.querySelectorAll('.sheet')) {
              sheet.style.zoom = String(layout.scale)
            }
          }
        })()`,
      )
      .then(() => {
        if (gen !== layoutGen || preview.webContents.isDestroyed()) return
        fittedPage = pageKey
        refitPreview()
      })
      .catch(() => undefined)
  }

  /// sheet/width modes use the preview-measured scale so print matches the
  /// left pane; scale mode keeps the user’s percent.
  const resolvePrintScale = async (choice: PrintDialogChoice): Promise<number> => {
    if (choice.fitMode === 'scale') return choice.scale
    if (preview.webContents.isDestroyed()) return choice.scale
    try {
      const scale = (await preview.webContents.executeJavaScript(
        'typeof window.__contentScale === "number" ? window.__contentScale : null',
      )) as number | null
      if (typeof scale === 'number' && Number.isFinite(scale) && scale >= 0.1 && scale <= 2) {
        return Math.round(scale * 1000) / 1000
      }
    } catch {
      /* preview torn down */
    }
    return choice.scale
  }

  let settled = false
  const choice = await new Promise<PrintDialogChoice | { error: string } | null>((resolve) => {
    let names = new Set<string>()
    let onChoice: (event: IpcMainEvent, payload: unknown) => void = () => {}
    let onLayout: (event: IpcMainEvent, payload: unknown) => void = () => {}
    const finish = (value: PrintDialogChoice | { error: string } | null) => {
      if (settled) return
      settled = true
      ipcMain.removeListener(channel, onChoice)
      ipcMain.removeListener(layoutChannel, onLayout)
      resolve(value)
    }
    onChoice = (_event, payload) => {
      if (payload == null) {
        finish(null)
        return
      }
      const parsed = parsePrintDialogChoice(payload, names)
      if (!parsed) {
        finish({ error: label('appPrintFailed') })
        return
      }
      finish(parsed)
    }
    onLayout = (_event, payload) => {
      const next = parsePreviewLayout(payload)
      if (!next) return
      void applyPreviewLayout(next)
    }
    ipcMain.on(channel, onChoice)
    ipcMain.on(layoutChannel, onLayout)
    win.once('closed', () => finish(null))
    void (async () => {
      try {
        const listed = await preview.webContents.getPrintersAsync()
        const printers: ListedPrinter[] = listed.map((printer) => ({
          name: printer.name,
          displayName: printer.displayName || printer.name,
        }))
        names = new Set(printers.map((printer) => printer.name))
        await writeFile(panelPath, panelHtml(printers, args.settings), 'utf8')
        await preview.webContents.loadFile(args.htmlPath)
        await preview.webContents.executeJavaScript(PREVIEW_FIT_SCRIPT)
        await applyPreviewLayout(args.settings)
        await win.loadFile(panelPath)
        layout()
        if (!win.isDestroyed()) {
          win.show()
          win.focus()
          layout()
        }
      } catch (error: unknown) {
        finish({ error: error instanceof Error ? error.message : String(error) })
      }
    })()
  })

  if (!choice || 'error' in choice) {
    if (!win.isDestroyed()) win.destroy()
    return choice && 'error' in choice ? { ok: false, error: choice.error } : { ok: false }
  }

  try {
    const scaledChoice = {
      ...choice,
      scale: await resolvePrintScale(choice),
    }
    if (isSaveAsPdfDevice(choice.deviceName)) {
      // Pick the path while the print window is still open. Destroying the
      // modal first made macOS drop the save panel with no UI (looked like
      // "Print to PDF" did nothing).
      const dialogParent = !win.isDestroyed() ? win : parent
      const selection = await showSaveDialogWithMemory(dialog, dialogParent, {
        title: label('appExportPdfTitle'),
        defaultPath: args.fileName,
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
      })
      if (!win.isDestroyed()) await closePrintWindow(win)
      if (selection.canceled || !selection.filePath) return { ok: false }
      return await savePdfFromPrintDialog({
        htmlPath: args.htmlPath,
        workDir: args.workDir,
        outPath: selection.filePath,
        choice: scaledChoice,
      })
    }

    if (!win.isDestroyed()) await closePrintWindow(win)
    return await printFromPrintDialog({
      htmlPath: args.htmlPath,
      workDir: args.workDir,
      choice: scaledChoice,
    })
  } catch (error: unknown) {
    if (!win.isDestroyed()) win.destroy()
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}
