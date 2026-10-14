/// One print window: page preview on the left, printer and page setup on the
/// right. Printing goes straight to the chosen printer, so the system print
/// panel does not open as a second window.

import { randomUUID } from 'node:crypto'
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import { BrowserWindow, dialog, ipcMain } from 'electron'

/// Bundled into out/main/chunks so the packaged app can copy them into the
/// temp preview folder (pdfjs-dist is not shipped as a loose node_modules tree).
import pdfJsAsset from 'pdfjs-dist/legacy/build/pdf.min.mjs?asset'
import pdfWorkerAsset from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?asset'

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
  applyPaperHandlingPages,
  parsePrintDialogChoice,
  printOptionsFromDialog,
  pdfPageIndicesInOrder,
  printToPdfOptionsFromDialog,
  requestFromPrintDialog,
  resolvePrintPageList,
} from './print-options'
import type { IpcMainEvent, PrinterInfo, WebContents, WebContentsPrintOptions } from 'electron'
import type { WorkbookPrintResult } from '../shared/desktop-api'
import type { PreviewLayout } from '../shared/preview-layout'
import type { PrintContentScope, WorkbookExportPdfRequest } from '../shared/desktop-api'
import type { PrintDialogChoice, PrintDialogDefaults } from './print-options'

export interface PrintContentMeta {
  scope: PrintContentScope
  sheetCount: number
  hasSelection: boolean
}

export type PrintContentRebuild = (scope: PrintContentScope) => Promise<
  | { request: WorkbookExportPdfRequest; settings: PrintDialogDefaults }
  | { error: string }
>

export type PrintPdfRenderer = (
  contents: WebContents,
  request: WorkbookExportPdfRequest,
) => Promise<Buffer>

/// Left preview width at the default window size (stays ~this when the
/// dialog is opened; sidebar is the rest).
const PREVIEW_PX = 300
const SIDEBAR_PX = 420
const DIALOG_WIDTH = PREVIEW_PX + SIDEBAR_PX
const DIALOG_MIN_WIDTH = PREVIEW_PX + 360

interface ListedPrinter {
  name: string
  displayName: string
  /// CUPS / OS “Location”.
  location: string
  /// Make-and-model or printer description.
  model: string
}

function optionString(options: Record<string, unknown>, key: string): string {
  const value = options[key]
  return typeof value === 'string' ? value.trim() : ''
}

function listedPrinterOf(printer: PrinterInfo): ListedPrinter {
  const options = (printer.options ?? {}) as Record<string, unknown>
  const location =
    optionString(options, 'printer-location') ||
    optionString(options, 'printer_location') ||
    optionString(options, 'location')
  const model =
    optionString(options, 'printer-make-and-model') ||
    optionString(options, 'printer_make_and_model') ||
    (printer.description || '').trim()
  return {
    name: printer.name,
    displayName: printer.displayName || printer.name,
    location,
    model,
  }
}

const PAPER_DIMS_MM: Record<(typeof NAMED_PAGE_SIZES)[number], { w: number; h: number }> = {
  A3: { w: 297, h: 420 },
  A4: { w: 210, h: 297 },
  A5: { w: 148, h: 210 },
  Letter: { w: 216, h: 279 },
  Legal: { w: 216, h: 356 },
  Tabloid: { w: 279, h: 432 },
}

type DialogLabelKey =
  | 'appPrintDialogTitle'
  | 'appPrintPrinter'
  | 'appPrintSaveAsPdf'
  | 'appPrintPdf'
  | 'appPrintCopies'
  | 'appPrintNow'
  | 'appPrintNoPrinter'
  | 'appPrintProperties'
  | 'appPrintFailed'
  | 'appExportPdfTitle'
  | 'appCancel'
  | 'appPrintPageOf'
  | 'appPrintRange'
  | 'appPrintRangeAll'
  | 'appPrintRangeCurrent'
  | 'appPrintRangeCustom'
  | 'appPrintRangeHint'
  | 'appPrintPageCount'
  | 'appPrintPageFrom'
  | 'appPrintPageTo'
  | 'appPrintColorMode'
  | 'appPrintColor'
  | 'appPrintGrayscale'
  | 'appPrintSides'
  | 'appPrintOneSided'
  | 'appPrintLongEdge'
  | 'appPrintShortEdge'
  | 'appSizeLabel'
  | 'appOrientationLabel'
  | 'appPortrait'
  | 'appLandscape'
  | 'appPrintCustomPaper'
  | 'appPrintMm'
  | 'appWidth'
  | 'appHeight'
  | 'appPrintSheetSection'
  | 'appPrintContent'
  | 'appPrintActiveSheet'
  | 'appPrintEntireWorkbook'
  | 'appPrintSelection'
  | 'appPrintAlign'
  | 'appPrintCenterHorizontalShort'
  | 'appPrintCenterVerticalShort'
  | 'appPrintScaleRatio'
  | 'appPrintNormalSize'
  | 'appPrintAdjustTo'
  | 'appPrintFitSheetOnePage'
  | 'appPrintFitColumnsOnePage'
  | 'appPrintLayoutSection'
  | 'appPrintPagesPerSheet'
  | 'appPrintLayoutDirection'
  | 'appPrintMarginsInches'
  | 'appPrintMarginTop'
  | 'appPrintMarginBottom'
  | 'appPrintMarginLeft'
  | 'appPrintMarginRight'
  | 'appPrintBorder'
  | 'appPrintBorderNone'
  | 'appPrintBorderHairline'
  | 'appPrintReverseOrientation'
  | 'appPrintFlipHorizontal'
  | 'appPrintPaperHandlingSection'
  | 'appPrintCollateSheets'
  | 'appPrintSheetsToPrint'
  | 'appPrintSheetsAll'
  | 'appPrintSheetsOdd'
  | 'appPrintSheetsEven'
  | 'appPrintPageOrder'
  | 'appPrintPageOrderAuto'
  | 'appPrintPageOrderNormal'
  | 'appPrintPageOrderReverse'
  | 'appPrintScaleToFit'
  | 'appPrintDestinationPaperSize'
  | 'appPrintScaleDownOnly'
  | 'appPrintPrinterInfoSection'
  | 'appPrintPrinterName'
  | 'appPrintPrinterLocation'
  | 'appPrintPrinterModel'
  | 'appPrintValueEmpty'

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

function namedPaperOptions(selected: string): string {
  const unit = label('appPrintMm')
  return NAMED_PAGE_SIZES.map((name) => {
    const { w, h } = PAPER_DIMS_MM[name]
    const text = `${name} ${w} × ${h} ${unit}`
    return `<option value="${name}"${selected === name ? ' selected' : ''}>${escapeHtml(text)}</option>`
  }).join('')
}

function paperOptions(defaults: PrintDialogDefaults): string {
  const custom = typeof defaults.pageSize !== 'string'
  const selected = typeof defaults.pageSize === 'string' ? defaults.pageSize : ''
  const customOption = `<option value="custom"${custom ? ' selected' : ''}>${escapeHtml(label('appPrintCustomPaper'))}</option>`
  return namedPaperOptions(selected) + customOption
}

function destinationPaperOptions(defaults: PrintDialogDefaults): string {
  const selected = typeof defaults.pageSize === 'string' ? defaults.pageSize : 'A4'
  return namedPaperOptions(selected)
}

/// Right sidebar UI mirrors the macOS system print dialog (label/control rows,
/// orientation toggles, PDF + Cancel + Print footer). Left preview is an
/// same-origin iframe so Layout toggles call `__applyPrintLayout` directly.
function panelHtml(
  printers: readonly ListedPrinter[],
  defaults: PrintDialogDefaults,
  pageCount: number,
  printContent: PrintContentMeta,
  documentMode = false,
  previewSrc = 'preview-frame.html',
): string {
  const data = JSON.stringify({
    printers,
    defaults,
    pageCount,
    printContent,
    documentMode,
    previewSrc,
  }).replace(/</g, '\\u003c')
  const none = printers.length === 0
  const custom = typeof defaults.pageSize === 'object' ? defaults.pageSize : { width: 8.5, height: 11 }
  const maxPage = Math.max(1, pageCount)
  const scope = printContent.scope
  const selectionDisabled = printContent.hasSelection ? '' : ' disabled'
  const sheetSection = documentMode
    ? ''
    : `
  <section class="section collapsible" id="sheet-section">
    <button type="button" class="section-head" data-toggle="sheet-section" aria-expanded="true">
      <span>${escapeHtml(label('appPrintSheetSection'))}</span>
      <svg class="chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M3 4.5 L6 7.5 L9 4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </button>
    <div class="section-body">
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintContent'))}</div>
        <div class="s-ctl">
          <select id="print-content">
            <option value="active"${scope === 'active' ? ' selected' : ''}>${escapeHtml(label('appPrintActiveSheet'))}</option>
            <option value="workbook"${scope === 'workbook' ? ' selected' : ''}>${escapeHtml(label('appPrintEntireWorkbook'))}</option>
            <option value="selection"${scope === 'selection' ? ' selected' : ''}${selectionDisabled}>${escapeHtml(label('appPrintSelection'))}</option>
          </select>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintAlign'))}</div>
        <div class="s-ctl">
          <label class="radio">
            <input id="center-h" type="checkbox"${defaults.horizontalCentered ? ' checked' : ''}>
            ${escapeHtml(label('appPrintCenterHorizontalShort'))}
          </label>
          <label class="radio">
            <input id="center-v" type="checkbox"${defaults.verticalCentered ? ' checked' : ''}>
            ${escapeHtml(label('appPrintCenterVerticalShort'))}
          </label>
        </div>
      </div>
      <div class="s-row scale-row">
        <label class="scale-line">
          <input type="radio" name="scale-mode" id="scale-mode-pct" value="scale"${defaults.fitMode === 'scale' ? ' checked' : ''}>
          ${escapeHtml(label('appPrintScaleRatio'))}
          <input id="scale" type="number" min="10" max="200" step="1" value="${Math.round(defaults.scale * 100)}"${defaults.fitMode === 'scale' ? '' : ' disabled'}>
          <span>${escapeHtml(label('appPrintNormalSize'))}</span>
        </label>
      </div>
      <div class="s-row scale-row">
        <label class="scale-line">
          <input type="radio" name="scale-mode" id="scale-mode-fit" value="fit"${defaults.fitMode === 'scale' ? '' : ' checked'}>
          ${escapeHtml(label('appPrintAdjustTo'))}
          <select id="fit-adjust"${defaults.fitMode === 'scale' ? ' disabled' : ''}>
            <option value="sheet"${defaults.fitMode !== 'width' ? ' selected' : ''}>${escapeHtml(label('appPrintFitSheetOnePage'))}</option>
            <option value="width"${defaults.fitMode === 'width' ? ' selected' : ''}>${escapeHtml(label('appPrintFitColumnsOnePage'))}</option>
          </select>
        </label>
      </div>
    </div>
  </section>`
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
  html, body {
    margin: 0;
    height: 100%;
    overflow: hidden;
    display: flex;
    flex-direction: row;
    background: Canvas;
    color: CanvasText;
  }
  [hidden] { display: none !important; }
  /* Same-origin iframe: Layout pane calls contentWindow.__applyPrintLayout
     directly (no cross-WebContents IPC / z-order). Flex height is explicit. */
  #preview-pane {
    flex: 1 1 auto;
    min-width: 0;
    min-height: 0;
    height: 100%;
    background: Canvas;
  }
  #preview-frame {
    display: block;
    width: 100%;
    height: 100%;
    border: 0;
    background: Canvas;
  }
  aside {
    flex: 0 0 ${SIDEBAR_PX}px;
    width: ${SIDEBAR_PX}px;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    min-height: 0;
    height: 100%;
    overflow: hidden;
    background: Canvas;
    border-left: 1px solid color-mix(in srgb, CanvasText 12%, transparent);
    font: 13px/1.35 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    -webkit-font-smoothing: antialiased;
  }
  .fields {
    flex: 1 1 auto;
    min-height: 0;
    overflow-x: hidden;
    overflow-y: auto;
    overscroll-behavior: contain;
    display: flex;
    flex-direction: column;
    gap: 12px;
    padding: 16px 14px 14px;
    scrollbar-width: thin;
    scrollbar-color: color-mix(in srgb, CanvasText 26%, transparent) transparent;
  }
  .fields::-webkit-scrollbar { width: 8px; }
  .fields::-webkit-scrollbar-thumb {
    background: color-mix(in srgb, CanvasText 26%, transparent);
    border-radius: 4px;
  }
  .title {
    margin: 0 2px 2px;
    font-size: 16px;
    font-weight: 650;
    letter-spacing: -0.02em;
    line-height: 1.25;
  }
  .form-row {
    display: grid;
    grid-template-columns: 88px 1fr;
    align-items: center;
    gap: 10px;
    min-height: 28px;
  }
  .form-row > .lbl {
    text-align: right;
    font-size: 12px;
    color: color-mix(in srgb, CanvasText 68%, transparent);
    line-height: 1.2;
  }
  .form-row > .ctl { min-width: 0; }
  .form-row.top { align-items: start; }
  .form-row.top > .lbl { padding-top: 6px; }
  select, input[type="number"], input[type="text"] {
    font: inherit;
    width: 100%;
    box-sizing: border-box;
    padding: 5px 9px;
    border-radius: 7px;
    border: 1px solid color-mix(in srgb, CanvasText 18%, transparent);
    background: Field;
    color: FieldText;
    outline: none;
    transition: border-color 0.12s ease, box-shadow 0.12s ease;
  }
  select:hover, input[type="number"]:hover, input[type="text"]:hover {
    border-color: color-mix(in srgb, CanvasText 30%, transparent);
  }
  select:focus-visible, input[type="number"]:focus-visible, input[type="text"]:focus-visible {
    border-color: AccentColor;
    box-shadow: 0 0 0 3px color-mix(in srgb, AccentColor 28%, transparent);
  }
  input[type="number"].short { width: 4.2em; }
  .inline {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 6px 8px;
  }
  .pages {
    display: flex;
    flex-direction: column;
    gap: 7px;
  }
  .radio {
    display: flex;
    align-items: center;
    gap: 7px;
    cursor: pointer;
    font-size: 12.5px;
  }
  .radio input { margin: 0; accent-color: AccentColor; }
  .orient {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
  }
  .orient label {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 5px;
    padding: 9px 4px 7px;
    border: 1px solid color-mix(in srgb, CanvasText 16%, transparent);
    border-radius: 9px;
    cursor: pointer;
    background: Field;
    color: FieldText;
    font-size: 11px;
    transition: border-color 0.12s ease, background 0.12s ease, box-shadow 0.12s ease;
  }
  .orient label:hover {
    border-color: color-mix(in srgb, CanvasText 28%, transparent);
  }
  .orient label:has(input:checked) {
    border-color: AccentColor;
    background: color-mix(in srgb, AccentColor 14%, Field);
    box-shadow: 0 0 0 1px color-mix(in srgb, AccentColor 35%, transparent);
  }
  .orient input { position: absolute; opacity: 0; pointer-events: none; }
  .orient svg { display: block; }
  .pair { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .margin-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
    width: 100%;
  }
  .margin-grid label {
    display: flex;
    flex-direction: column;
    gap: 3px;
    font-size: 11px;
    color: color-mix(in srgb, CanvasText 72%, transparent);
    min-width: 0;
  }
  .margin-grid input[type="number"] {
    width: 100%;
    box-sizing: border-box;
  }
  .hint {
    margin: 0 2px;
    font-size: 11px;
    line-height: 1.4;
    color: color-mix(in srgb, CanvasText 62%, transparent);
  }
  .actions {
    flex: none;
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 12px 14px 14px;
    border-top: 1px solid color-mix(in srgb, CanvasText 12%, transparent);
    background: color-mix(in srgb, CanvasText 3%, Canvas);
  }
  .actions-end { display: flex; gap: 8px; margin-left: auto; }
  button {
    font: inherit;
    font-size: 13px;
    padding: 6px 14px;
    border-radius: 7px;
    border: 1px solid color-mix(in srgb, CanvasText 18%, transparent);
    background: Field;
    color: FieldText;
    cursor: pointer;
    transition: background 0.12s ease, border-color 0.12s ease, opacity 0.12s ease;
  }
  button:hover:not(:disabled) {
    background: color-mix(in srgb, CanvasText 6%, Field);
    border-color: color-mix(in srgb, CanvasText 28%, transparent);
  }
  button:focus-visible {
    outline: none;
    box-shadow: 0 0 0 3px color-mix(in srgb, AccentColor 28%, transparent);
  }
  button.primary {
    background: AccentColor;
    color: AccentColorText;
    border-color: transparent;
    min-width: 76px;
    font-weight: 600;
  }
  button.primary:hover:not(:disabled) {
    background: color-mix(in srgb, CanvasText 12%, AccentColor);
    border-color: transparent;
  }
  button:disabled { opacity: 0.45; cursor: default; }
  button.pdf { font-weight: 550; }
  .section {
    margin-top: 0;
    border-radius: 11px;
    border: 1px solid color-mix(in srgb, CanvasText 12%, transparent);
    background: color-mix(in srgb, CanvasText 4%, Canvas);
    box-shadow: 0 1px 2px color-mix(in srgb, CanvasText 6%, transparent);
    flex: none;
    overflow: hidden;
  }
  .section > .section-head {
    margin: 0;
    width: 100%;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    padding: 10px 12px 9px;
    font: inherit;
    font-size: 12px;
    font-weight: 650;
    letter-spacing: 0.01em;
    color: inherit;
    background: transparent;
    border: 0;
    border-bottom: 1px solid color-mix(in srgb, CanvasText 10%, transparent);
    border-radius: 0;
    cursor: pointer;
    text-align: left;
    box-sizing: border-box;
    transition: background 0.12s ease;
  }
  .section > .section-head:hover {
    background: color-mix(in srgb, CanvasText 5%, transparent);
  }
  .section .chev {
    flex: none;
    width: 12px;
    height: 12px;
    opacity: 0.55;
    transition: transform 0.16s ease, opacity 0.12s ease;
  }
  .section > .section-head:hover .chev { opacity: 0.8; }
  .section.collapsed > .section-head {
    border-bottom-color: transparent;
  }
  .section.collapsed > .section-head .chev {
    transform: rotate(-90deg);
  }
  .section.collapsed > .section-body {
    display: none;
  }
  .section .s-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 9px 12px;
    border-bottom: 1px solid color-mix(in srgb, CanvasText 8%, transparent);
    min-height: 36px;
    box-sizing: border-box;
  }
  .section .s-row:last-child { border-bottom: 0; }
  .section .s-row.top { align-items: flex-start; }
  .section .s-lbl {
    flex: 1 1 auto;
    min-width: 0;
    max-width: 58%;
    font-size: 12px;
    line-height: 1.35;
    color: color-mix(in srgb, CanvasText 76%, transparent);
    white-space: normal;
    overflow-wrap: break-word;
  }
  .section .s-ctl {
    flex: 0 1 auto;
    min-width: 0;
    max-width: 55%;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 8px;
    flex-wrap: wrap;
  }
  .section .s-ctl.col {
    flex-direction: column;
    align-items: stretch;
    gap: 7px;
    max-width: 68%;
    flex: 1 1 auto;
  }
  .section .s-ctl:has(.orient),
  .section .s-ctl:has(.dir-group) {
    max-width: 68%;
    flex: 1 1 auto;
  }
  .section .s-ctl select {
    max-width: 100%;
    width: auto;
    min-width: 0;
  }
  .section .s-row.scale-row {
    justify-content: flex-start;
  }
  .scale-line {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 6px;
    flex-wrap: nowrap;
    width: 100%;
    min-width: 0;
    font-size: 12px;
    cursor: pointer;
  }
  .scale-line input[type="number"] {
    width: 5.5em;
    min-width: 5.5em;
    flex: none;
    padding-inline: 6px 4px;
    appearance: textfield;
    -moz-appearance: textfield;
  }
  .scale-line input[type="number"]::-webkit-outer-spin-button,
  .scale-line input[type="number"]::-webkit-inner-spin-button {
    margin: 0;
  }
  .scale-line select {
    flex: 1 1 auto;
    width: auto;
    min-width: 0;
    max-width: none;
  }
  .dir-group {
    display: inline-flex;
    border: 1px solid color-mix(in srgb, CanvasText 16%, transparent);
    border-radius: 8px;
    overflow: hidden;
    background: Field;
  }
  .dir-group label {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 34px;
    height: 28px;
    cursor: pointer;
    border-right: 1px solid color-mix(in srgb, CanvasText 12%, transparent);
    transition: background 0.12s ease;
  }
  .dir-group label:last-child { border-right: 0; }
  .dir-group label:hover {
    background: color-mix(in srgb, CanvasText 5%, Field);
  }
  .dir-group label:has(input:checked) {
    background: color-mix(in srgb, AccentColor 20%, Field);
  }
  .dir-group input { position: absolute; opacity: 0; pointer-events: none; }
  .dir-group svg { display: block; opacity: 0.85; }
  .toggle {
    appearance: none;
    width: 40px;
    height: 22px;
    border-radius: 11px;
    background: color-mix(in srgb, CanvasText 26%, transparent);
    border: 0;
    position: relative;
    cursor: pointer;
    flex: none;
    transition: background 0.14s ease;
  }
  .toggle::after {
    content: '';
    position: absolute;
    top: 2px;
    left: 2px;
    width: 18px;
    height: 18px;
    border-radius: 50%;
    background: Canvas;
    box-shadow: 0 1px 2px color-mix(in srgb, CanvasText 25%, transparent);
    transition: left 0.14s ease;
  }
  .toggle:checked { background: AccentColor; }
  .toggle:checked::after { left: 20px; }
  .toggle:disabled { opacity: 0.45; cursor: default; }
  .section .s-row.indent .s-lbl { padding-left: 18px; }
  .section .s-row.dimmed { opacity: 0.45; }
  .section .s-row.dimmed select,
  .section .s-row.dimmed .toggle { pointer-events: none; }
  .section .s-val {
    flex: 1 1 auto;
    min-width: 0;
    text-align: right;
    font-size: 12px;
    color: color-mix(in srgb, CanvasText 88%, transparent);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
</style>
</head>
<body>
<div id="preview-pane">
  <iframe id="preview-frame" title="Print preview" src="${escapeHtml(previewSrc)}"></iframe>
</div>
<aside>
<div class="fields">
  <h1 class="title">${escapeHtml(label('appPrintDialogTitle'))}</h1>

  <section class="section collapsible" id="general-section">
    <button type="button" class="section-head" data-toggle="general-section" aria-expanded="true">
      <span>${escapeHtml(label('appPrintProperties'))}</span>
      <svg class="chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M3 4.5 L6 7.5 L9 4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </button>
    <div class="section-body">
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintPrinter'))}</div>
        <div class="s-ctl"><select id="printer"${none ? ' disabled' : ''}></select></div>
      </div>
      <div class="s-row" id="printer-copies">
        <div class="s-lbl">${escapeHtml(label('appPrintCopies'))}</div>
        <div class="s-ctl">
          <input id="copies" class="short" type="number" min="1" max="99" value="1">
        </div>
      </div>
      <div class="s-row top">
        <div class="s-lbl">${escapeHtml(label('appPrintRange'))}</div>
        <div class="s-ctl col">
          <label class="radio">
            <input type="radio" name="range" id="range-all" value="all" checked>
            ${escapeHtml(label('appPrintRangeAll'))}
          </label>
          <label class="radio">
            <input type="radio" name="range" id="range-current" value="current">
            ${escapeHtml(label('appPrintRangeCurrent'))}
          </label>
          <label class="radio">
            <input type="radio" name="range" id="range-fromto" value="fromto">
            <span class="inline">
              ${escapeHtml(label('appPrintPageFrom'))}
              <input id="page-from" class="short" type="number" min="1" max="${maxPage}" value="1" disabled>
              ${escapeHtml(label('appPrintPageTo'))}
              <input id="page-to" class="short" type="number" min="1" max="${maxPage}" value="${maxPage}" disabled>
            </span>
          </label>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appSizeLabel'))}</div>
        <div class="s-ctl"><select id="paper">${paperOptions(defaults)}</select></div>
      </div>
      <div class="s-row" id="custom-paper-row"${typeof defaults.pageSize === 'string' ? ' hidden' : ''}>
        <div class="s-lbl"></div>
        <div class="s-ctl pair" id="custom-paper">
          <input id="paper-width" type="number" min="1" max="100" step="0.01" value="${custom.width}" aria-label="${escapeHtml(label('appWidth'))}">
          <input id="paper-height" type="number" min="1" max="100" step="0.01" value="${custom.height}" aria-label="${escapeHtml(label('appHeight'))}">
        </div>
      </div>
      <div class="s-row top">
        <div class="s-lbl">${escapeHtml(label('appOrientationLabel'))}</div>
        <div class="s-ctl">
          <div class="orient">
            <label>
              <input type="radio" name="orientation" id="orient-portrait" value="portrait"${defaults.landscape ? '' : ' checked'}>
              <svg width="26" height="34" viewBox="0 0 26 34" aria-hidden="true">
                <rect x="3" y="2" width="20" height="30" rx="2" fill="none" stroke="currentColor" stroke-width="1.5"/>
                <line x1="7" y1="9" x2="19" y2="9" stroke="currentColor" stroke-width="1"/>
                <line x1="7" y1="13" x2="19" y2="13" stroke="currentColor" stroke-width="1"/>
                <line x1="7" y1="17" x2="15" y2="17" stroke="currentColor" stroke-width="1"/>
              </svg>
              ${escapeHtml(label('appPortrait'))}
            </label>
            <label>
              <input type="radio" name="orientation" id="orient-landscape" value="landscape"${defaults.landscape ? ' checked' : ''}>
              <svg width="34" height="26" viewBox="0 0 34 26" aria-hidden="true">
                <rect x="2" y="3" width="30" height="20" rx="2" fill="none" stroke="currentColor" stroke-width="1.5"/>
                <line x1="7" y1="9" x2="27" y2="9" stroke="currentColor" stroke-width="1"/>
                <line x1="7" y1="13" x2="27" y2="13" stroke="currentColor" stroke-width="1"/>
                <line x1="7" y1="17" x2="19" y2="17" stroke="currentColor" stroke-width="1"/>
              </svg>
              ${escapeHtml(label('appLandscape'))}
            </label>
          </div>
        </div>
      </div>
      <div class="s-row" id="color-field">
        <div class="s-lbl">${escapeHtml(label('appPrintColorMode'))}</div>
        <div class="s-ctl">
          <select id="color">
            <option value="1">${escapeHtml(label('appPrintColor'))}</option>
            <option value="0">${escapeHtml(label('appPrintGrayscale'))}</option>
          </select>
        </div>
      </div>
      <div class="s-row" id="duplex-field">
        <div class="s-lbl">${escapeHtml(label('appPrintSides'))}</div>
        <div class="s-ctl">
          <select id="duplex">
            <option value="simplex">${escapeHtml(label('appPrintOneSided'))}</option>
            <option value="longEdge">${escapeHtml(label('appPrintLongEdge'))}</option>
            <option value="shortEdge">${escapeHtml(label('appPrintShortEdge'))}</option>
          </select>
        </div>
      </div>
    </div>
  </section>

  ${sheetSection}

  <section class="section collapsible" id="layout-section">
    <button type="button" class="section-head" data-toggle="layout-section" aria-expanded="true">
      <span>${escapeHtml(label('appPrintLayoutSection'))}</span>
      <svg class="chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M3 4.5 L6 7.5 L9 4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </button>
    <div class="section-body">
      <div class="s-row top">
        <div class="s-lbl">${escapeHtml(label('appPrintMarginsInches'))}</div>
        <div class="s-ctl margin-grid">
          <label>
            <span>${escapeHtml(label('appPrintMarginTop'))}</span>
            <input id="margin-top" type="number" min="0" max="3" step="0.05" value="${defaults.margins.top}">
          </label>
          <label>
            <span>${escapeHtml(label('appPrintMarginBottom'))}</span>
            <input id="margin-bottom" type="number" min="0" max="3" step="0.05" value="${defaults.margins.bottom}">
          </label>
          <label>
            <span>${escapeHtml(label('appPrintMarginLeft'))}</span>
            <input id="margin-left" type="number" min="0" max="3" step="0.05" value="${defaults.margins.left}">
          </label>
          <label>
            <span>${escapeHtml(label('appPrintMarginRight'))}</span>
            <input id="margin-right" type="number" min="0" max="3" step="0.05" value="${defaults.margins.right}">
          </label>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintPagesPerSheet'))}</div>
        <div class="s-ctl">
          <input id="pages-per-sheet" class="short" type="number" min="1" max="16" value="1">
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintLayoutDirection'))}</div>
        <div class="s-ctl">
          <div class="dir-group" role="radiogroup" aria-label="${escapeHtml(label('appPrintLayoutDirection'))}">
            <label title="Z">
              <input type="radio" name="layout-dir" value="ltr-ttb" checked>
              <svg width="18" height="14" viewBox="0 0 18 14" aria-hidden="true"><path d="M2 3h10M12 3l-3-2M12 3l-3 2M6 11h10M16 11l-3-2M16 11l-3 2M4 3v8" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
            </label>
            <label title="Z-rev">
              <input type="radio" name="layout-dir" value="rtl-ttb">
              <svg width="18" height="14" viewBox="0 0 18 14" aria-hidden="true"><path d="M16 3H6M6 3l3-2M6 3l3 2M12 11H2M2 11l3-2M2 11l3 2M14 3v8" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
            </label>
            <label title="N">
              <input type="radio" name="layout-dir" value="ttb-ltr">
              <svg width="18" height="14" viewBox="0 0 18 14" aria-hidden="true"><path d="M4 2v8M4 10l-2-3M4 10l2-3M14 4v8M14 12l-2-3M14 12l2-3M4 3h10" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
            </label>
            <label title="N-rev">
              <input type="radio" name="layout-dir" value="ttb-rtl">
              <svg width="18" height="14" viewBox="0 0 18 14" aria-hidden="true"><path d="M14 2v8M14 10l-2-3M14 10l2-3M4 4v8M4 12l-2-3M4 12l2-3M14 3H4" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
            </label>
          </div>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintBorder'))}</div>
        <div class="s-ctl">
          <select id="page-border">
            <option value="none" selected>${escapeHtml(label('appPrintBorderNone'))}</option>
            <option value="hairline">${escapeHtml(label('appPrintBorderHairline'))}</option>
          </select>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintReverseOrientation'))}</div>
        <div class="s-ctl">
          <input id="reverse-orient" class="toggle" type="checkbox">
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintFlipHorizontal'))}</div>
        <div class="s-ctl">
          <input id="flip-h" class="toggle" type="checkbox">
        </div>
      </div>
    </div>
  </section>

  <section class="section collapsible" id="paper-handling-section">
    <button type="button" class="section-head" data-toggle="paper-handling-section" aria-expanded="true">
      <span>${escapeHtml(label('appPrintPaperHandlingSection'))}</span>
      <svg class="chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M3 4.5 L6 7.5 L9 4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </button>
    <div class="section-body">
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintCollateSheets'))}</div>
        <div class="s-ctl">
          <input id="collate" class="toggle" type="checkbox" checked>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintSheetsToPrint'))}</div>
        <div class="s-ctl">
          <select id="sheets-to-print">
            <option value="all" selected>${escapeHtml(label('appPrintSheetsAll'))}</option>
            <option value="odd">${escapeHtml(label('appPrintSheetsOdd'))}</option>
            <option value="even">${escapeHtml(label('appPrintSheetsEven'))}</option>
          </select>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintPageOrder'))}</div>
        <div class="s-ctl">
          <select id="page-order">
            <option value="automatic" selected>${escapeHtml(label('appPrintPageOrderAuto'))}</option>
            <option value="normal">${escapeHtml(label('appPrintPageOrderNormal'))}</option>
            <option value="reverse">${escapeHtml(label('appPrintPageOrderReverse'))}</option>
          </select>
        </div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintScaleToFit'))}</div>
        <div class="s-ctl">
          <input id="scale-to-fit" class="toggle" type="checkbox">
        </div>
      </div>
      <div class="s-row indent dimmed" id="dest-paper-row">
        <div class="s-lbl">${escapeHtml(label('appPrintDestinationPaperSize'))}</div>
        <div class="s-ctl">
          <select id="dest-paper" disabled>${destinationPaperOptions(defaults)}</select>
        </div>
      </div>
      <div class="s-row indent dimmed" id="scale-down-row">
        <div class="s-lbl">${escapeHtml(label('appPrintScaleDownOnly'))}</div>
        <div class="s-ctl">
          <input id="scale-down-only" class="toggle" type="checkbox" disabled>
        </div>
      </div>
    </div>
  </section>

  <section class="section collapsible" id="printer-info-section">
    <button type="button" class="section-head" data-toggle="printer-info-section" aria-expanded="true">
      <span>${escapeHtml(label('appPrintPrinterInfoSection'))}</span>
      <svg class="chev" viewBox="0 0 12 12" aria-hidden="true"><path d="M3 4.5 L6 7.5 L9 4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
    </button>
    <div class="section-body">
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintPrinterName'))}</div>
        <div class="s-val" id="printer-info-name">${escapeHtml(label('appPrintValueEmpty'))}</div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintPrinterLocation'))}</div>
        <div class="s-val" id="printer-info-location">${escapeHtml(label('appPrintValueEmpty'))}</div>
      </div>
      <div class="s-row">
        <div class="s-lbl">${escapeHtml(label('appPrintPrinterModel'))}</div>
        <div class="s-val" id="printer-info-model">${escapeHtml(label('appPrintValueEmpty'))}</div>
      </div>
    </div>
  </section>

  ${none ? `<p class="hint" id="no-printer">${escapeHtml(label('appPrintNoPrinter'))}</p>` : ''}
</div>
  <div class="actions">
    <button id="save-pdf" class="pdf" type="button" title="${escapeHtml(label('appPrintSaveAsPdf'))}">PDF ▾</button>
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
  if (data.printers.length > 0) select.value = data.printers[0].name
  const emptyVal = ${JSON.stringify(label('appPrintValueEmpty'))}
  const infoName = document.getElementById('printer-info-name')
  const infoLocation = document.getElementById('printer-info-location')
  const infoModel = document.getElementById('printer-info-model')
  const printerOf = (name) => data.printers.find((p) => p.name === name) || null
  const syncPrinterMeta = () => {
    const printer = printerOf(select.value)
    infoName.textContent = (printer && (printer.displayName || printer.name)) || emptyVal
    infoLocation.textContent = (printer && printer.location) || emptyVal
    infoModel.textContent = (printer && printer.model) || emptyVal
  }
  select.addEventListener('change', syncPrinterMeta)
  syncPrinterMeta()
  const printContent = document.getElementById('print-content')
  let contentBusy = false
  if (printContent) printContent.addEventListener('change', async () => {
    if (!window.printDialog || !window.printDialog.setContent || contentBusy) return
    const scope = printContent.value
    contentBusy = true
    printContent.disabled = true
    try {
      const result = await window.printDialog.setContent(scope)
      if (result && result.error) {
        printContent.value = data.printContent.scope
        return
      }
      if (result && result.ok) {
        data.printContent.scope = scope
        if (typeof result.pageCount === 'number' && result.pageCount >= 1) {
          data.pageCount = result.pageCount
          pageFrom.max = String(result.pageCount)
          pageTo.max = String(result.pageCount)
          if (Number(pageTo.value) > result.pageCount) pageTo.value = String(result.pageCount)
        }
      }
    } finally {
      contentBusy = false
      printContent.disabled = false
    }
  })
  const rangeAll = document.getElementById('range-all')
  const rangeCurrent = document.getElementById('range-current')
  const rangeFromTo = document.getElementById('range-fromto')
  const pageFrom = document.getElementById('page-from')
  const pageTo = document.getElementById('page-to')
  const syncRange = () => {
    const on = rangeFromTo.checked
    pageFrom.disabled = !on
    pageTo.disabled = !on
  }
  const compactRange = (pages) => {
    const s = [...new Set(pages)].filter((n) => n >= 1).sort((a, b) => a - b)
    if (!s.length) return ''
    const parts = []
    let start = s[0], prev = s[0]
    for (let i = 1; i < s.length; i++) {
      if (s[i] === prev + 1) { prev = s[i]; continue }
      parts.push(start === prev ? String(start) : start + '-' + prev)
      start = prev = s[i]
    }
    parts.push(start === prev ? String(start) : start + '-' + prev)
    return parts.join(',')
  }
  let ignorePick = false
  const previewWin = () => {
    const frame = document.getElementById('preview-frame')
    return frame && frame.contentWindow
  }
  const pickedFromPreview = () => {
    const w = previewWin()
    if (!w || typeof w.__getSelectedPages !== 'function') return null
    return w.__getSelectedPages()
  }
  const syncPrintEnabled = () => {
    const picked = pickedFromPreview()
    const none = Array.isArray(picked) && picked.length === 0
    const printEl = document.getElementById('print')
    const saveEl = document.getElementById('save-pdf')
    if (printEl) printEl.disabled = data.printers.length === 0 || none
    if (saveEl) saveEl.disabled = none
  }
  const pushRangeToPreview = () => {
    if (ignorePick) return
    const w = previewWin()
    if (!w) return
    if (rangeAll.checked) {
      if (typeof w.__selectAllPages === 'function') w.__selectAllPages()
      return
    }
    if (rangeCurrent.checked) {
      const cur = typeof w.__getCurrentPage === 'function' ? w.__getCurrentPage() : 1
      if (typeof w.__setSelectedPages === 'function') w.__setSelectedPages([cur])
      return
    }
    let from = Math.round(num('page-from')) || 1
    let to = Math.round(num('page-to')) || from
    if (from > to) { const s = from; from = to; to = s }
    const list = []
    for (let i = from; i <= to; i++) list.push(i)
    if (typeof w.__setSelectedPages === 'function') w.__setSelectedPages(list)
  }
  rangeAll.onchange = () => { syncRange(); pushRangeToPreview(); syncPrintEnabled() }
  rangeCurrent.onchange = () => { syncRange(); pushRangeToPreview(); syncPrintEnabled() }
  rangeFromTo.onchange = () => { syncRange(); pushRangeToPreview(); syncPrintEnabled() }
  pageFrom.onfocus = () => { rangeFromTo.checked = true; syncRange() }
  pageTo.onfocus = () => { rangeFromTo.checked = true; syncRange() }
  pageFrom.addEventListener('change', () => { pushRangeToPreview(); syncPrintEnabled() })
  pageTo.addEventListener('change', () => { pushRangeToPreview(); syncPrintEnabled() })
  window.addEventListener('message', (event) => {
    if (!event.data || event.data.type !== 'print-page-selection') return
    const pages = Array.isArray(event.data.pages) ? event.data.pages : []
    const total = Math.round(Number(event.data.total)) || 0
    if (total >= 1) {
      data.pageCount = total
      pageFrom.max = String(total)
      pageTo.max = String(total)
    }
    ignorePick = true
    if (pages.length === 0) {
      rangeFromTo.checked = true
    } else if (total >= 1 && pages.length === total) {
      rangeAll.checked = true
    } else {
      rangeFromTo.checked = true
      const sorted = [...pages].sort((a, b) => a - b)
      const contig = sorted.every((n, i) => i === 0 || n === sorted[i - 1] + 1)
      if (contig && sorted.length) {
        pageFrom.value = String(sorted[0])
        pageTo.value = String(sorted[sorted.length - 1])
      }
    }
    syncRange()
    ignorePick = false
    syncPrintEnabled()
  })
  syncRange()
  const paper = document.getElementById('paper')
  const customPaperRow = document.getElementById('custom-paper-row')
  paper.onchange = () => { customPaperRow.hidden = paper.value !== 'custom' }
  const noPrinter = document.getElementById('no-printer')
  const printBtn = document.getElementById('print')
  const savePdfBtn = document.getElementById('save-pdf')
  const syncDestination = () => {
    printBtn.disabled = data.printers.length === 0
    if (noPrinter) noPrinter.hidden = data.printers.length > 0
  }
  syncDestination()
  syncPrintEnabled()
  for (const head of document.querySelectorAll('.section.collapsible > .section-head')) {
    head.addEventListener('click', () => {
      const section = head.closest('.section')
      const collapsed = section.classList.toggle('collapsed')
      head.setAttribute('aria-expanded', collapsed ? 'false' : 'true')
    })
  }
  // Keep wheel/trackpad scrolling inside the sidebar so the left preview
  // iframe never receives scroll-chained events from this pane.
  document.querySelector('.fields').addEventListener(
    'wheel',
    (event) => { event.stopPropagation() },
    { passive: true },
  )
  const num = (id) => {
    const el = document.getElementById(id)
    return el ? Number(el.value) : 0
  }
  const scaleInput = document.getElementById('scale')
  const fitAdjust = document.getElementById('fit-adjust')
  const scaleModePct = document.getElementById('scale-mode-pct')
  const scaleModeFit = document.getElementById('scale-mode-fit')
  const syncScaleMode = () => {
    if (!scaleModePct || !scaleInput || !fitAdjust) return
    const pct = scaleModePct.checked
    scaleInput.disabled = !pct
    fitAdjust.disabled = pct
  }
  if (scaleModePct && scaleModeFit) {
    scaleModePct.onchange = syncScaleMode
    scaleModeFit.onchange = syncScaleMode
    syncScaleMode()
  }
  const scaleToFit = document.getElementById('scale-to-fit')
  const destPaper = document.getElementById('dest-paper')
  const scaleDownOnly = document.getElementById('scale-down-only')
  const destPaperRow = document.getElementById('dest-paper-row')
  const scaleDownRow = document.getElementById('scale-down-row')
  const syncScaleToFit = () => {
    const on = scaleToFit.checked
    destPaper.disabled = !on
    scaleDownOnly.disabled = !on
    destPaperRow.classList.toggle('dimmed', !on)
    scaleDownRow.classList.toggle('dimmed', !on)
  }
  scaleToFit.onchange = () => { syncScaleToFit(); publishLayout() }
  syncScaleToFit()
  const rangeMode = () => {
    if (rangeCurrent.checked) return 'current'
    if (rangeFromTo.checked) return 'custom'
    return 'all'
  }
  const customRangeValue = () => {
    if (!rangeFromTo.checked) return ''
    let from = Math.round(num('page-from')) || 1
    let to = Math.round(num('page-to')) || from
    if (from > to) { const s = from; from = to; to = s }
    return from + '-' + to
  }
  const currentFitMode = () => {
    if (!scaleModePct || scaleModePct.checked) return 'scale'
    return fitAdjust && fitAdjust.value === 'width' ? 'width' : 'sheet'
  }
  const layoutDirection = () => {
    const el = document.querySelector('input[name="layout-dir"]:checked')
    return el ? el.value : 'ltr-ttb'
  }
  const sourcePageSize = () => paper.value === 'custom'
    ? { width: num('paper-width'), height: num('paper-height') }
    : paper.value
  const clampMargin = (value) => {
    const n = Number(value)
    if (!Number.isFinite(n)) return 0
    return Math.min(3, Math.max(0, Math.round(n * 100) / 100))
  }
  const currentMargins = () => {
    const d = data.defaults.margins || { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 }
    const read = (id, fallback) => {
      const el = document.getElementById(id)
      return el ? clampMargin(el.value) : clampMargin(fallback)
    }
    return {
      top: read('margin-top', d.top),
      bottom: read('margin-bottom', d.bottom),
      left: read('margin-left', d.left),
      right: read('margin-right', d.right),
    }
  }
  const currentLayout = () => {
    const d = data.defaults
    const centerH = document.getElementById('center-h')
    const centerV = document.getElementById('center-v')
    const reverseOrient = document.getElementById('reverse-orient')
    const flipH = document.getElementById('flip-h')
    const pageBorder = document.getElementById('page-border')
    return {
      landscape: document.getElementById('orient-landscape').checked,
      pageSize: scaleToFit.checked ? destPaper.value : sourcePageSize(),
      margins: currentMargins(),
      scale: (Math.round(num('scale')) || Math.round(d.scale * 100) || 100) / 100,
      fitMode: currentFitMode(),
      horizontalCentered: centerH ? centerH.checked : !!d.horizontalCentered,
      verticalCentered: centerV ? centerV.checked : !!d.verticalCentered,
      reverseOrientation: reverseOrient ? reverseOrient.checked : false,
      flipHorizontal: flipH ? flipH.checked : false,
      pageBorder: pageBorder && pageBorder.value === 'hairline' ? 'hairline' : 'none',
      pagesPerSheet: Math.min(16, Math.max(1, Math.round(num('pages-per-sheet')) || 1)),
      layoutDirection: layoutDirection(),
    }
  }
  /// Mirror of previewLayoutPatch — runs in the sidebar so Layout toggles
  /// update the iframe without waiting on main-process IPC.
  const PAPER_IN = {
    Letter: { width: 8.5, height: 11 },
    Tabloid: { width: 11, height: 17 },
    Legal: { width: 8.5, height: 14 },
    A3: { width: 11.69, height: 16.54 },
    A4: { width: 8.27, height: 11.69 },
    A5: { width: 5.83, height: 8.27 },
  }
  const nupGrid = (n) => {
    const v = Math.min(16, Math.max(1, Math.round(n) || 1))
    const hit = [1, 2, 4, 6, 9, 16].find((x) => x >= v) || 16
    if (hit === 2) return { cols: 2, rows: 1 }
    if (hit === 4) return { cols: 2, rows: 2 }
    if (hit === 6) return { cols: 3, rows: 2 }
    if (hit === 9) return { cols: 3, rows: 3 }
    if (hit === 16) return { cols: 4, rows: 4 }
    return { cols: 1, rows: 1 }
  }
  const layoutPatch = (L) => {
    const paper = typeof L.pageSize === 'string'
      ? (PAPER_IN[L.pageSize] || PAPER_IN.A4)
      : L.pageSize
    const across = L.landscape ? paper.height : paper.width
    const down = L.landscape ? paper.width : paper.height
    const round3 = (x) => Math.round(x * 1000) / 1000
    const grid = nupGrid(L.pagesPerSheet || 1)
    const dir = L.layoutDirection === 'rtl-ttb' || L.layoutDirection === 'ttb-ltr' || L.layoutDirection === 'ttb-rtl'
      ? L.layoutDirection : 'ltr-ttb'
    return {
      pageW: round3(across),
      pageH: round3(down),
      top: round3(L.margins.top),
      right: round3(L.margins.right),
      bottom: round3(L.margins.bottom),
      left: round3(L.margins.left),
      scale: round3(Math.min(2, Math.max(0.1, L.scale))),
      fitMode: L.fitMode,
      align: L.horizontalCentered ? 'center' : 'flex-start',
      justify: L.verticalCentered ? 'center' : 'flex-start',
      reverseOrientation: !!L.reverseOrientation,
      flipHorizontal: !!L.flipHorizontal,
      pageBorder: L.pageBorder === 'hairline' ? 'hairline' : 'none',
      pagesPerSheet: grid.cols * grid.rows,
      layoutDirection: dir,
      nupCols: grid.cols,
      nupRows: grid.rows,
    }
  }
  const previewFrame = document.getElementById('preview-frame')
  const applyToPreview = (patch) => {
    const w = previewFrame && previewFrame.contentWindow
    if (!w || typeof w.__applyPrintLayout !== 'function') return false
    try {
      w.__applyPrintLayout(patch)
      return true
    } catch {
      return false
    }
  }
  const publishLayout = () => {
    const payload = currentLayout()
    const patch = layoutPatch(payload)
    if (!applyToPreview(patch) && previewFrame) {
      const onLoad = () => { applyToPreview(patch) }
      previewFrame.addEventListener('load', onLoad, { once: true })
      // Module PDF preview may expose __applyPrintLayout slightly after load.
      let tries = 0
      const poll = setInterval(() => {
        tries += 1
        if (applyToPreview(patch) || tries > 80) clearInterval(poll)
      }, 50)
    }
    if (window.printDialog && window.printDialog.layout) void window.printDialog.layout(payload)
  }
  for (const id of [
    'paper', 'orient-portrait', 'orient-landscape', 'paper-width', 'paper-height',
    'scale', 'scale-mode-pct', 'scale-mode-fit', 'fit-adjust', 'center-h', 'center-v',
    'dest-paper', 'reverse-orient', 'flip-h', 'page-border', 'pages-per-sheet',
    'margin-top', 'margin-bottom', 'margin-left', 'margin-right',
  ]) {
    const el = document.getElementById(id)
    if (!el) continue
    el.addEventListener('input', publishLayout)
    el.addEventListener('change', publishLayout)
  }
  for (const el of document.querySelectorAll('input[name="layout-dir"]')) {
    el.addEventListener('change', publishLayout)
  }
  const submitChoice = (deviceName) => {
    if (!window.printDialog || !deviceName) return
    printBtn.disabled = true
    savePdfBtn.disabled = true
    document.getElementById('cancel').disabled = true
    const layout = currentLayout()
    const picked = pickedFromPreview()
    let nextRangeMode = rangeMode()
    let nextCustom = customRangeValue()
    if (picked) {
      const total = (previewWin() && previewWin().__getLogicalPageCount) ? previewWin().__getLogicalPageCount() : data.pageCount
      if (picked.length === 0) {
        printBtn.disabled = false
        savePdfBtn.disabled = false
        document.getElementById('cancel').disabled = false
        return
      }
      if (picked.length === total) {
        nextRangeMode = 'all'
        nextCustom = ''
      } else {
        nextRangeMode = 'custom'
        nextCustom = compactRange(picked)
      }
    }
    window.printDialog.submit({
      deviceName,
      copies: Math.round(num('copies')) || 1,
      collate: document.getElementById('collate').checked,
      color: document.getElementById('color').value === '1',
      duplexMode: document.getElementById('duplex').value,
      rangeMode: nextRangeMode,
      customRange: nextCustom,
      landscape: layout.landscape,
      pageSize: sourcePageSize(),
      margins: layout.margins,
      scale: layout.scale,
      fitMode: layout.fitMode,
      horizontalCentered: layout.horizontalCentered,
      verticalCentered: layout.verticalCentered,
      pagesPerSheet: Math.round(num('pages-per-sheet')) || 1,
      layoutDirection: layoutDirection(),
      pageBorder: document.getElementById('page-border').value,
      reverseOrientation: document.getElementById('reverse-orient').checked,
      flipHorizontal: document.getElementById('flip-h').checked,
      sheetsToPrint: document.getElementById('sheets-to-print').value,
      pageOrder: document.getElementById('page-order').value,
      scaleToFit: scaleToFit.checked,
      destinationPageSize: destPaper.value,
      scaleDownOnly: scaleDownOnly.checked,
    })
  }
  document.getElementById('cancel').onclick = () => {
    if (window.printDialog) window.printDialog.submit(null)
  }
  savePdfBtn.onclick = () => submitChoice(SAVE_AS_PDF)
  printBtn.onclick = () => submitChoice(select.value)
  // Sync left preview to the sidebar's initial layout (Word PDF.js included).
  publishLayout()
</script>
</body>
</html>`
}

function preloadSource(channel: string, layoutChannel: string, contentChannel: string): string {
  return `const { contextBridge, ipcRenderer } = require('electron')
contextBridge.exposeInMainWorld('printDialog', {
  submit: (choice) => ipcRenderer.send(${JSON.stringify(channel)}, choice),
  layout: (layout) => ipcRenderer.invoke(${JSON.stringify(layoutChannel)}, layout),
  setContent: (scope) => ipcRenderer.invoke(${JSON.stringify(contentChannel)}, scope),
})
`
}

/// Builds stacked page cards for screen preview (vertical pagination), then
/// fits page width into the pane. Pinch zoom only changes the pane scale;
/// `__contentScale` is what print/PDF must use. Source HTML on disk is
/// unchanged — print still paginates via Chromium.
///
/// Performance with many pages:
/// - Slots are lightweight paper placeholders (fixed pane-scaled size).
/// - Heavy sheet DOM mounts only after scroll settles, and only for the
///   visible page(s).
/// - Main process freezes live visible pages to bitmaps; later scrolls paint
///   cheap images instead of re-cloning spreadsheet tables.
const PREVIEW_FIT_SCRIPT = `(() => {
  if (window.__fitPreview) return
  let userZoom = 1
  let baseScale = 1
  let fitMode = 'sheet'
  let paneScale = 1
  let pageWpx = 0
  let pageHpx = 0
  let cardMetrics = []
  let pageSpecs = []
  let labelRaf = 0
  let paneRaf = 0
  let mountTimer = 0
  let lastLabel = ''
  let built = false
  let scrolling = false
  let reverseOrientation = false
  let flipHorizontal = false
  let pageBorder = 'none'
  let nupCols = 1
  let nupRows = 1
  let layoutDirection = 'ltr-ttb'
  let selectedPages = null
  let logicalPageCount = 0
  const VIRTUALIZE_AFTER = 1
  const VISIBLE_BUFFER = 0
  const MOUNT_IDLE_MS = 70
  const pageOf = ${JSON.stringify(label('appPrintPageOf'))}
  const nupCellIndex = (fillIndex, cols, rows, direction) => {
    const total = cols * rows
    if (fillIndex < 0 || fillIndex >= total || cols < 1 || rows < 1) return fillIndex
    let row = 0
    let col = 0
    if (direction === 'rtl-ttb') {
      row = Math.floor(fillIndex / cols)
      col = cols - 1 - (fillIndex % cols)
    } else if (direction === 'ttb-ltr') {
      col = Math.floor(fillIndex / rows)
      row = fillIndex % rows
    } else if (direction === 'ttb-rtl') {
      col = cols - 1 - Math.floor(fillIndex / rows)
      row = fillIndex % rows
    } else {
      row = Math.floor(fillIndex / cols)
      col = fillIndex % cols
    }
    return row * cols + col
  }
  const allPageNums = (n) => {
    const out = []
    for (let i = 1; i <= n; i++) out.push(i)
    return out
  }
  const pageIsOn = (n) => !selectedPages || selectedPages.has(n)
  const pruneSelection = (total) => {
    logicalPageCount = Math.max(0, total)
    if (!selectedPages) return
    const next = new Set()
    for (const n of selectedPages) if (n >= 1 && n <= logicalPageCount) next.add(n)
    selectedPages = next.size === logicalPageCount ? null : next
  }
  const notifySelection = () => {
    try {
      window.parent.postMessage({
        type: 'print-page-selection',
        pages: window.__getSelectedPages(),
        total: logicalPageCount,
      }, '*')
    } catch { /* parent gone */ }
  }
  const paintPicks = () => {
    for (const wrap of pages.querySelectorAll('.preview-sheet')) {
      let anyOn = false
      let anyOff = false
      for (const btn of wrap.querySelectorAll('.preview-page-pick')) {
        const n = Number(btn.dataset.page)
        const on = pageIsOn(n)
        btn.setAttribute('aria-pressed', on ? 'true' : 'false')
        if (on) anyOn = true
        else anyOff = true
      }
      wrap.classList.toggle('is-off', anyOff && !anyOn)
    }
  }
  const togglePage = (n) => {
    if (!(n >= 1 && n <= logicalPageCount)) return
    if (!selectedPages) selectedPages = new Set(allPageNums(logicalPageCount))
    if (selectedPages.has(n)) selectedPages.delete(n)
    else selectedPages.add(n)
    if (selectedPages.size === logicalPageCount) selectedPages = null
    paintPicks()
    notifySelection()
  }
  const makePickBtn = (pageNo) => {
    const btn = document.createElement('button')
    btn.type = 'button'
    btn.className = 'preview-page-pick'
    btn.dataset.page = String(pageNo)
    btn.textContent = String(pageNo)
    btn.setAttribute('aria-pressed', 'true')
    btn.addEventListener('click', (event) => {
      event.preventDefault()
      event.stopPropagation()
      togglePage(pageNo)
    })
    return btn
  }
  const makePickRow = (spec) => {
    const row = document.createElement('div')
    row.className = 'preview-page-picks'
    if (spec && spec.nup) {
      for (const p of spec.pages) row.appendChild(makePickBtn(p.pageNo))
    } else {
      row.appendChild(makePickBtn(spec.pageNo))
    }
    return row
  }
  const slotOf = (wrap) =>
    wrap && wrap.querySelector ? wrap.querySelector('.preview-card-slot') || wrap : wrap
  window.__getSelectedPages = () =>
    selectedPages ? [...selectedPages].sort((a, b) => a - b) : allPageNums(logicalPageCount)
  window.__getLogicalPageCount = () => logicalPageCount
  window.__selectAllPages = () => {
    selectedPages = null
    paintPicks()
    return true
  }
  window.__setSelectedPages = (list) => {
    selectedPages = new Set()
    if (Array.isArray(list)) {
      for (const raw of list) {
        const n = Math.round(Number(raw))
        if (n >= 1 && n <= logicalPageCount) selectedPages.add(n)
      }
    }
    if (selectedPages.size === logicalPageCount) selectedPages = null
    paintPicks()
    return true
  }
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
  const measureHost = document.createElement('div')
  measureHost.setAttribute('aria-hidden', 'true')
  measureHost.style.cssText =
    'position:fixed;left:-12000px;top:0;width:auto;height:auto;overflow:visible;visibility:hidden;pointer-events:none;'
  document.body.appendChild(measureHost)

  const applyPageChrome = (area) => {
    const root = document.documentElement
    const pageW = root.style.getPropertyValue('--page-w') || getComputedStyle(root).getPropertyValue('--page-w')
    const pageH = root.style.getPropertyValue('--page-h') || getComputedStyle(root).getPropertyValue('--page-h')
    const top = root.style.getPropertyValue('--m-top') || getComputedStyle(root).getPropertyValue('--m-top')
    const right = root.style.getPropertyValue('--m-right') || getComputedStyle(root).getPropertyValue('--m-right')
    const bottom = root.style.getPropertyValue('--m-bottom') || getComputedStyle(root).getPropertyValue('--m-bottom')
    const left = root.style.getPropertyValue('--m-left') || getComputedStyle(root).getPropertyValue('--m-left')
    const align = (root.style.getPropertyValue('--align-x') || getComputedStyle(root).getPropertyValue('--align-x') || 'flex-start').trim() || 'flex-start'
    const justify = (root.style.getPropertyValue('--align-y') || getComputedStyle(root).getPropertyValue('--align-y') || 'flex-start').trim() || 'flex-start'
    area.style.boxSizing = 'border-box'
    area.style.display = 'flex'
    area.style.flexDirection = 'column'
    if (pageW) area.style.width = pageW.trim() + 'in'
    if (pageH) area.style.height = pageH.trim() + 'in'
    area.style.padding = [top, right, bottom, left].map((v) => (v || '0').trim() + 'in').join(' ')
    // Alignment lives on .layout-fx so reverse/flip can wrap content without
    // fighting the pane-fit scale on .area.
    area.style.justifyContent = 'stretch'
    area.style.alignItems = 'stretch'
    area.style.overflow = 'hidden'
    if (pageBorder === 'hairline') {
      area.style.outline = '1px solid #000'
      area.style.outlineOffset = '-1px'
    } else {
      area.style.outline = ''
      area.style.outlineOffset = ''
    }
    let fx = area.querySelector(':scope > .layout-fx')
    if (!fx) {
      fx = document.createElement('div')
      fx.className = 'layout-fx'
      while (area.firstChild) fx.appendChild(area.firstChild)
      area.appendChild(fx)
    }
    fx.style.boxSizing = 'border-box'
    fx.style.flex = '1 1 auto'
    fx.style.minHeight = '0'
    fx.style.width = '100%'
    fx.style.height = '100%'
    fx.style.display = 'flex'
    fx.style.flexDirection = 'column'
    fx.style.justifyContent = justify
    fx.style.alignItems = align
    fx.style.overflow = 'hidden'
    const parts = []
    if (reverseOrientation) parts.push('rotate(180deg)')
    if (flipHorizontal) parts.push('scaleX(-1)')
    fx.style.transform = parts.length ? parts.join(' ') : ''
    fx.style.transformOrigin = 'center center'
  }

  const applySheetScale = (sheet, zoom) => {
    sheet.style.zoom = ''
    sheet.style.transform = zoom === 1 ? '' : 'scale(' + zoom + ')'
    sheet.style.transformOrigin = 'top left'
    sheet.style.position = ''
    sheet.style.top = ''
    sheet.style.flex = 'none'
    sheet.style.width = 'max-content'
    sheet.style.minWidth = 'max-content'
  }

  const applyAreaPaneTransform = (area) => {
    area.style.transform = 'scale(' + paneScale + ')'
    area.style.transformOrigin = 'top left'
  }

  const cacheCardMetrics = () => {
    // Prefer arithmetic metrics when every slot shares one size — avoids
    // reading layout for hundreds of nodes on every scroll tick.
    const first = pages.firstElementChild
    if (!first || !pageSpecs.length) {
      cardMetrics = []
      return
    }
    const h = first.offsetHeight
    const gap = 14
    const base = first.offsetTop
    if (h > 0) {
      cardMetrics = pageSpecs.map((_, i) => ({
        top: base + i * (h + gap),
        height: h,
      }))
      return
    }
    cardMetrics = [...pages.children].map((wrap) => ({
      top: wrap.offsetTop,
      height: wrap.offsetHeight,
    }))
  }

  const updatePageLabel = () => {
    const total = Math.max(1, cardMetrics.length || pageSpecs.length)
    let current = 1
    if (cardMetrics.length > 0) {
      const mid = scroll.scrollTop + scroll.clientHeight * 0.35
      // Binary search — many-page workbooks otherwise walk hundreds of slots.
      let lo = 0
      let hi = cardMetrics.length - 1
      while (lo <= hi) {
        const midIdx = (lo + hi) >> 1
        const top = cardMetrics[midIdx].top
        const bottom = top + cardMetrics[midIdx].height
        if (mid < top) hi = midIdx - 1
        else if (mid >= bottom) lo = midIdx + 1
        else {
          current = midIdx + 1
          break
        }
      }
      if (lo > hi) current = Math.min(total, Math.max(1, lo + 1))
    }
    const text = pageOf
      .replace('{current}', String(current))
      .replace('{total}', String(total))
    if (text !== lastLabel) {
      lastLabel = text
      pageLabel.textContent = text
    }
  }

  const schedulePageLabel = () => {
    if (labelRaf) return
    labelRaf = requestAnimationFrame(() => {
      labelRaf = 0
      updatePageLabel()
    })
  }

  const visibleRange = () => {
    const total = pages.childElementCount
    if (total === 0) return { first: 0, last: -1 }
    if (total <= VIRTUALIZE_AFTER) return { first: 0, last: total - 1 }
    if (!cardMetrics.length) cacheCardMetrics()
    const top = scroll.scrollTop
    const bottom = top + scroll.clientHeight
    let first = 0
    let last = total - 1
    for (let i = 0; i < cardMetrics.length; i++) {
      if (cardMetrics[i].top + cardMetrics[i].height >= top) {
        first = i
        break
      }
    }
    for (let i = first; i < cardMetrics.length; i++) {
      if (cardMetrics[i].top <= bottom) last = i
      else break
    }
    return {
      first: Math.max(0, first - VISIBLE_BUFFER),
      last: Math.min(total - 1, last + VISIBLE_BUFFER),
    }
  }

  const mountLogicalPage = (spec, nested) => {
    const source = sources[spec.sourceIndex]
    if (!source) return null
    const savedReverse = reverseOrientation
    const savedFlip = flipHorizontal
    const savedBorder = pageBorder
    if (nested) {
      // Sheet-level Layout toggles wrap the whole N-up card, not each cell.
      reverseOrientation = false
      flipHorizontal = false
      pageBorder = 'none'
    }
    const card = source.cloneNode(true)
    applyPageChrome(card)
    if (nested) {
      reverseOrientation = savedReverse
      flipHorizontal = savedFlip
      pageBorder = savedBorder
    }
    const cardSheet = card.querySelector('.sheet')
    if (!cardSheet) return card
    applySheetScale(cardSheet, spec.zoom)
    if (spec.pageCount > 1) {
      const fx = card.querySelector(':scope > .layout-fx')
      if (fx) {
        fx.style.justifyContent = 'flex-start'
        fx.style.alignItems = spec.alignX
      }
      const clip = document.createElement('div')
      clip.className = 'page-clip'
      clip.style.display = 'flex'
      clip.style.justifyContent = spec.alignX
      clip.style.alignItems = 'flex-start'
      cardSheet.style.position = 'relative'
      cardSheet.style.top =
        spec.pageIndex === 0 ? '0' : String(-spec.pageIndex * spec.ah / spec.zoom) + 'px'
      cardSheet.replaceWith(clip)
      clip.appendChild(cardSheet)
    }
    return card
  }

  const mountCard = (spec) => {
    if (spec && spec.nup) {
      const sheet = document.createElement('div')
      sheet.className = 'area'
      applyPageChrome(sheet)
      const fx = sheet.querySelector(':scope > .layout-fx')
      if (!fx) {
        applyAreaPaneTransform(sheet)
        return sheet
      }
      fx.style.display = 'grid'
      fx.style.gridTemplateColumns = 'repeat(' + spec.cols + ', 1fr)'
      fx.style.gridTemplateRows = 'repeat(' + spec.rows + ', 1fr)'
      fx.style.gap = '1px'
      fx.style.background = 'color-mix(in srgb, CanvasText 18%, transparent)'
      fx.style.justifyContent = 'stretch'
      fx.style.alignItems = 'stretch'
      const cellCount = spec.cols * spec.rows
      const cells = []
      for (let i = 0; i < cellCount; i++) {
        const cell = document.createElement('div')
        cell.style.overflow = 'hidden'
        cell.style.position = 'relative'
        cell.style.background = '#fff'
        cells.push(cell)
        fx.appendChild(cell)
      }
      const fit = Math.min(1 / spec.cols, 1 / spec.rows)
      for (let i = 0; i < spec.pages.length; i++) {
        const cell = cells[nupCellIndex(i, spec.cols, spec.rows, spec.dir)]
        if (!cell) continue
        const mini = mountLogicalPage(spec.pages[i], true)
        if (!mini) continue
        // Nested page chrome is already full-paper; scale into the N-up cell.
        mini.style.position = 'absolute'
        mini.style.top = '0'
        mini.style.left = '0'
        mini.style.transform = 'scale(' + fit + ')'
        mini.style.transformOrigin = 'top left'
        cell.appendChild(mini)
      }
      applyAreaPaneTransform(sheet)
      return sheet
    }
    const card = mountLogicalPage(spec, false)
    if (!card) return null
    applyAreaPaneTransform(card)
    return card
  }

  const clearBitmaps = () => {
    for (const wrap of pages.children) {
      const slot = slotOf(wrap)
      if (slot.dataset.frozen !== '1') continue
      slot.replaceChildren()
      slot.dataset.frozen = '0'
      slot.dataset.live = '0'
      slot.dataset.empty = '1'
    }
  }

  /// Mount live DOM only for visible unfrozen slots. Never touch frozen images.
  /// During active scrolling, skip mounts so the compositor stays free.
  const syncVisibleCards = (force) => {
    if (!built || pageSpecs.length === 0) return
    if (scrolling && !force) return
    const { first, last } = visibleRange()
    const virtualize = pageSpecs.length > VIRTUALIZE_AFTER
    for (let i = 0; i < pages.childElementCount; i++) {
      const slot = slotOf(pages.children[i])
      if (slot.dataset.frozen === '1') {
        slot.dataset.empty = '0'
        slot.dataset.live = '0'
        continue
      }
      const want = !virtualize || (i >= first && i <= last)
      if (want) {
        if (slot.dataset.live !== '1') {
          const card = mountCard(pageSpecs[i])
          if (card) {
            slot.replaceChildren(card)
            slot.dataset.empty = '0'
            slot.dataset.live = '1'
          }
        } else {
          const area = slot.querySelector('.area')
          if (area) applyAreaPaneTransform(area)
        }
      } else if (slot.dataset.live === '1') {
        // Drop live DOM off-screen; keep empty paper until frozen/revisit.
        slot.replaceChildren()
        slot.dataset.empty = '1'
        slot.dataset.live = '0'
      }
    }
  }

  const scheduleMount = () => {
    if (mountTimer) clearTimeout(mountTimer)
    mountTimer = setTimeout(() => {
      mountTimer = 0
      scrolling = false
      syncVisibleCards(true)
    }, MOUNT_IDLE_MS)
  }

  const onScroll = () => {
    scrolling = true
    schedulePageLabel()
    scheduleMount()
  }

  const rebuildPages = () => {
    if (mountTimer) {
      clearTimeout(mountTimer)
      mountTimer = 0
    }
    pages.innerHTML = ''
    pageSpecs = []
    built = false
    pageWpx = 0
    pageHpx = 0
    let contentScale = baseScale
    const logical = []
    measureHost.replaceChildren()
    for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex++) {
      const source = sources[sourceIndex]
      const probe = source.cloneNode(true)
      measureHost.appendChild(probe)
      applyPageChrome(probe)
      const sheet = probe.querySelector('.sheet')
      if (!sheet) {
        probe.remove()
        continue
      }
      applySheetScale(sheet, 1)
      const style = getComputedStyle(probe)
      const padX = (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0)
      const padY = (parseFloat(style.paddingTop) || 0) + (parseFloat(style.paddingBottom) || 0)
      pageWpx = Math.max(pageWpx, probe.offsetWidth)
      pageHpx = Math.max(pageHpx, probe.offsetHeight)
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
      const alignX = (getComputedStyle(document.documentElement)
        .getPropertyValue('--align-x') || 'flex-start').trim() || 'flex-start'
      probe.remove()
      for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
        logical.push({
          sourceIndex,
          pageIndex,
          pageCount,
          zoom,
          ah,
          alignX,
          pageNo: logical.length + 1,
        })
      }
    }
    const perSheet = Math.max(1, nupCols * nupRows)
    if (perSheet <= 1) {
      pageSpecs = logical
    } else {
      for (let i = 0; i < logical.length; i += perSheet) {
        pageSpecs.push({
          nup: true,
          cols: nupCols,
          rows: nupRows,
          dir: layoutDirection,
          pages: logical.slice(i, i + perSheet),
        })
      }
    }
    for (let i = 0; i < pageSpecs.length; i++) {
      const wrap = document.createElement('div')
      wrap.className = 'preview-sheet'
      const slot = document.createElement('div')
      slot.className = 'preview-card-slot'
      slot.dataset.empty = '1'
      slot.dataset.live = '0'
      slot.dataset.frozen = '0'
      wrap.appendChild(slot)
      wrap.appendChild(makePickRow(pageSpecs[i]))
      pages.appendChild(wrap)
    }
    pruneSelection(logical.length)
    paintPicks()
    notifySelection()
    window.__contentScale = contentScale
    window.__fitMode = fitMode
    built = pageSpecs.length > 0
  }

  const sizeSlots = () => {
    if (!pageWpx || !pageHpx) return false
    if (scroll.clientWidth < 40) return false
    const paneW = Math.max(scroll.clientWidth, fitHost.clientWidth, window.innerWidth || 0)
    const fitted = (paneW - 24) / pageWpx
    const next = Math.max(0.15, Math.min(4, fitted * userZoom))
    const scaleChanged = Math.abs(next - paneScale) >= 0.0005
    const needsSize = !slotOf(pages.firstElementChild)?.style.width
    if (!scaleChanged && !needsSize) {
      cacheCardMetrics()
      return true
    }
    if (scaleChanged) clearBitmaps()
    paneScale = next
    const slotW = Math.ceil(pageWpx * paneScale) + 'px'
    const slotH = Math.ceil(pageHpx * paneScale) + 'px'
    for (const wrap of pages.children) {
      const slot = slotOf(wrap)
      slot.style.width = slotW
      slot.style.height = slotH
      slot.style.containIntrinsicSize = slotW + ' ' + slotH
      const area = slot.querySelector('.area')
      if (area) applyAreaPaneTransform(area)
    }
    cacheCardMetrics()
    return true
  }

  const applyPaneScale = () => {
    if (!built) rebuildPages()
    if (!built) {
      cardMetrics = []
      updatePageLabel()
      return
    }
    sizeSlots()
    scrolling = false
    syncVisibleCards(true)
    updatePageLabel()
  }

  const schedulePaneScale = () => {
    if (paneRaf) return
    paneRaf = requestAnimationFrame(() => {
      paneRaf = 0
      applyPaneScale()
    })
  }

  const rebuildAndFit = () => {
    rebuildPages()
    applyPaneScale()
  }

  window.__fitPreview = schedulePaneScale
  window.__setPrintScale = (scale, mode) => {
    baseScale = typeof scale === 'number' && scale > 0 ? scale : 1
    fitMode = mode === 'sheet' || mode === 'width' ? mode : 'scale'
    rebuildAndFit()
  }
  /// Sidebar layout IPC: paper / margins / scale / center / N-up / flip / border.
  /// Clears frozen bitmaps so cards remount with the new chrome.
  window.__applyPrintLayout = (layout) => {
    if (!layout) return false
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
    reverseOrientation = !!layout.reverseOrientation
    flipHorizontal = !!layout.flipHorizontal
    pageBorder = layout.pageBorder === 'hairline' ? 'hairline' : 'none'
    nupCols = Math.max(1, Math.round(Number(layout.nupCols)) || 1)
    nupRows = Math.max(1, Math.round(Number(layout.nupRows)) || 1)
    layoutDirection =
      layout.layoutDirection === 'rtl-ttb' ||
      layout.layoutDirection === 'ttb-ltr' ||
      layout.layoutDirection === 'ttb-rtl'
        ? layout.layoutDirection
        : 'ltr-ttb'
    baseScale = typeof layout.scale === 'number' && layout.scale > 0 ? layout.scale : 1
    fitMode = layout.fitMode === 'sheet' || layout.fitMode === 'width' ? layout.fitMode : 'scale'
    clearBitmaps()
    rebuildAndFit()
    return true
  }
  /// Live+mostly-visible pages ready for main-process capturePage.
  window.__freezeCandidates = () => {
    if (!built || scrolling) return []
    const host = scroll.getBoundingClientRect()
    const out = []
    for (let i = 0; i < pages.childElementCount; i++) {
      const slot = slotOf(pages.children[i])
      if (slot.dataset.frozen === '1' || slot.dataset.live !== '1') continue
      const r = slot.getBoundingClientRect()
      const overlap = Math.min(r.bottom, host.bottom) - Math.max(r.top, host.top)
      if (overlap < r.height * 0.8 || r.width < 8 || r.height < 8) continue
      out.push({
        index: i,
        x: Math.round(r.x),
        y: Math.round(r.y),
        width: Math.round(r.width),
        height: Math.round(r.height),
      })
      if (out.length >= 2) break
    }
    return out
  }
  window.__getCurrentPage = () => {
    const total = Math.max(1, cardMetrics.length || pageSpecs.length)
    if (cardMetrics.length === 0) return 1
    const mid = scroll.scrollTop + scroll.clientHeight * 0.35
    let current = 1
    for (let i = 0; i < cardMetrics.length; i++) {
      const top = cardMetrics[i].top
      const bottom = top + cardMetrics[i].height
      if (mid >= top && mid < bottom) {
        current = i + 1
        break
      }
      if (mid >= bottom) current = i + 1
    }
    const spec = pageSpecs[Math.min(total, Math.max(1, current)) - 1]
    if (spec && spec.nup && spec.pages[0]) return spec.pages[0].pageNo
    if (spec && spec.pageNo) return spec.pageNo
    return Math.min(total, Math.max(1, current))
  }
  window.__applyPageBitmap = (index, dataUrl) => {
    const slot = slotOf(pages.children[index])
    if (!slot || typeof dataUrl !== 'string' || !dataUrl.startsWith('data:')) return false
    const img = document.createElement('img')
    img.className = 'preview-card-bitmap'
    img.alt = ''
    img.draggable = false
    img.decoding = 'async'
    img.src = dataUrl
    slot.replaceChildren(img)
    slot.dataset.frozen = '1'
    slot.dataset.live = '0'
    slot.dataset.empty = '0'
    return true
  }
  scroll.addEventListener('scroll', onScroll, { passive: true })
  window.addEventListener('resize', schedulePaneScale)
  window.addEventListener('wheel', (event) => {
    if (!event.ctrlKey) return
    event.preventDefault()
    userZoom = Math.max(0.2, Math.min(5, userZoom * Math.exp(-event.deltaY * 0.01)))
    schedulePaneScale()
  }, { passive: false })
  rebuildAndFit()
})()`

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
    | 'pageBorder'
    | 'reverseOrientation'
    | 'flipHorizontal'
    | 'scaleToFit'
    | 'destinationPageSize'
  >,
): Promise<void> {
  const layoutCss = previewLayoutCss({
    landscape: choice.landscape,
    pageSize: choice.scaleToFit ? choice.destinationPageSize : choice.pageSize,
    margins: choice.margins,
    scale: choice.scale,
    // Job window prints via scaleFactor; keep CSS on fixed scale.
    fitMode: 'scale',
    horizontalCentered: choice.horizontalCentered,
    verticalCentered: choice.verticalCentered,
  })
  const transforms: string[] = []
  if (choice.reverseOrientation) transforms.push('rotate(180deg)')
  if (choice.flipHorizontal) transforms.push('scaleX(-1)')
  const extras: string[] = []
  if (transforms.length > 0) {
    extras.push(
      `@media print { .area { transform: ${transforms.join(' ')}; transform-origin: center center; } }`,
    )
  }
  if (choice.pageBorder === 'hairline') {
    extras.push('@media print { .area { outline: 1px solid #000; outline-offset: -1px; } }')
  }
  const style = `${layoutCss}\n${extras.join('\n')}`
  const source = await readFile(sourcePath, 'utf8')
  const html = source.includes('</head>')
    ? source.replace('</head>', `<style id="go-preview-layout">${style}</style></head>`)
    : `${source}<style id="go-preview-layout">${style}</style>`
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
async function layoutPdfBytes(args: {
  htmlPath: string
  workDir: string
  fileName: string
  choice: PrintDialogChoice
  sourceRequest: WorkbookExportPdfRequest | undefined
  renderPdf: PrintPdfRenderer | undefined
}): Promise<Buffer> {
  const jobPath = join(args.workDir, args.fileName)
  await writeJobHtml(args.htmlPath, jobPath, args.choice)
  const hidden = await openJobWindow(jobPath)
  try {
    const pdf =
      args.renderPdf && args.sourceRequest
        ? await args.renderPdf(hidden.webContents, requestFromPrintDialog(args.sourceRequest, args.choice))
        : await hidden.webContents.printToPDF(printToPdfOptionsFromDialog(args.choice))
    if (!args.choice.pageList) return pdf
    return await subsetPdfPages(pdf, args.choice.pageList)
  } finally {
    if (!hidden.isDestroyed()) hidden.destroy()
  }
}

async function savePdfFromPrintDialog(args: {
  htmlPath: string
  workDir: string
  outPath: string
  choice: PrintDialogChoice
  sourceRequest: WorkbookExportPdfRequest | undefined
  renderPdf: PrintPdfRenderer | undefined
}): Promise<WorkbookPrintResult> {
  const pdf = await layoutPdfBytes({
    htmlPath: args.htmlPath,
    workDir: args.workDir,
    fileName: 'print-save.pdf.html',
    choice: args.choice,
    sourceRequest: args.sourceRequest,
    renderPdf: args.renderPdf,
  })
  await atomicWriteFile(args.outPath, pdf)
  return { ok: true, path: args.outPath }
}

async function closePrintWindow(win: BrowserWindow): Promise<void> {
  if (win.isDestroyed()) return
  await new Promise<void>((resolve) => {
    win.once('closed', () => resolve())
    win.destroy()
  })
}

async function readPdfPageCount(pdfPath: string): Promise<number> {
  const { PDFDocument } = await import('pdf-lib')
  const pdf = await PDFDocument.load(await readFile(pdfPath))
  return Math.max(1, pdf.getPageCount())
}

/// Document-mode left pane: PDF.js cards on `Canvas`. Loaded in the left
/// Word PDF preview document loaded in the print dialog iframe.
async function writePdfJsPreviewHtml(workDir: string, pdfPath: string): Promise<string> {
  const htmlPath = join(workDir, 'pdf-preview.html')
  const localPdf = join(workDir, 'preview.pdf')
  /// readFile+writeFile for pdf.js assets (they live inside app.asar).
  await writeFile(join(workDir, 'pdf.min.mjs'), await readFile(pdfJsAsset))
  await writeFile(join(workDir, 'pdf.worker.min.mjs'), await readFile(pdfWorkerAsset))
  /// Load via relative URL — embedding multi‑MB PDFs as base64 blocked the dialog.
  if (pdfPath !== localPdf) await copyFile(pdfPath, localPdf)
  await writeFile(
    htmlPath,
    `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
html, body {
  margin: 0;
  width: 100%;
  height: 100%;
  background: Canvas;
  color: CanvasText;
  overflow: hidden;
}
#scroll {
  box-sizing: border-box;
  width: 100%;
  height: 100%;
  overflow: auto;
  padding: 16px 12px 24px;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 16px;
  background: Canvas;
  scrollbar-width: thin;
  scrollbar-color: color-mix(in srgb, CanvasText 26%, transparent) transparent;
}
#scroll::-webkit-scrollbar { width: 8px; }
#scroll::-webkit-scrollbar-thumb {
  background: color-mix(in srgb, CanvasText 26%, transparent);
  border-radius: 4px;
}
.card {
  flex: none;
  box-sizing: border-box;
  background: #fff;
  box-shadow:
    0 0 0 1px color-mix(in srgb, CanvasText 14%, transparent),
    0 1px 2px color-mix(in srgb, CanvasText 10%, transparent),
    0 6px 18px color-mix(in srgb, CanvasText 12%, transparent);
  overflow: hidden;
  display: flex;
}
.card .nup-grid {
  flex: 1 1 auto;
  min-height: 0;
  width: 100%;
  height: 100%;
  display: grid;
  gap: 1px;
  background: color-mix(in srgb, CanvasText 18%, transparent);
}
.card .nup-cell {
  overflow: hidden;
  position: relative;
  background: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
}
.card .stage {
  flex: none;
  transform-origin: center center;
}
.card canvas { display: block; }
.card-wrap {
  flex: none;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 8px;
}
.card-wrap.is-off .card { opacity: 0.42; }
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
.err, .loading {
  margin: 24px;
  color: color-mix(in srgb, CanvasText 70%, transparent);
  font: 13px/1.4 system-ui, sans-serif;
}
</style></head><body>
<div id="scroll"><p class="loading">…</p></div>
<script type="module">
import * as pdfjs from './pdf.min.mjs';
pdfjs.GlobalWorkerOptions.workerSrc = './pdf.worker.min.mjs';
const scroll = document.getElementById('scroll');
const CSS_PX = 96;
const pageNodes = [];
let pendingLayout = null;
let userZoom = 1;
let fitRaf = 0;
let selectedPages = null;
let resolveReady;
let readySettled = false;
window.__pdfReady = new Promise((resolve) => { resolveReady = resolve; });
const markReady = () => {
  if (readySettled) return;
  readySettled = true;
  resolveReady();
};
const showPendingOrFallback = () => {
  if (pendingLayout) {
    applyPrintLayoutNow(pendingLayout);
  } else if (pageNodes.length > 0) {
    const first = pageNodes[0];
    applyPrintLayoutNow({
      pageW: first.pageWIn,
      pageH: first.pageHIn,
      top: 0, right: 0, bottom: 0, left: 0,
      scale: 1,
      fitMode: 'scale',
      align: 'center',
      justify: 'center',
      reverseOrientation: false,
      flipHorizontal: false,
      pageBorder: 'none',
      pagesPerSheet: 1,
      layoutDirection: 'ltr-ttb',
      nupCols: 1,
      nupRows: 1,
    });
  }
};
const nupCellIndex = (fillIndex, cols, rows, direction) => {
  const total = cols * rows;
  if (fillIndex < 0 || fillIndex >= total || cols < 1 || rows < 1) return fillIndex;
  let row = 0;
  let col = 0;
  if (direction === 'rtl-ttb') {
    row = Math.floor(fillIndex / cols);
    col = cols - 1 - (fillIndex % cols);
  } else if (direction === 'ttb-ltr') {
    col = Math.floor(fillIndex / rows);
    row = fillIndex % rows;
  } else if (direction === 'ttb-rtl') {
    col = cols - 1 - Math.floor(fillIndex / rows);
    row = fillIndex % rows;
  } else {
    row = Math.floor(fillIndex / cols);
    col = fillIndex % cols;
  }
  return row * cols + col;
};
const allPageNums = (n) => {
  const out = [];
  for (let i = 1; i <= n; i++) out.push(i);
  return out;
};
const pageIsOn = (n) => !selectedPages || selectedPages.has(n);
const notifySelection = () => {
  try {
    window.parent.postMessage({
      type: 'print-page-selection',
      pages: window.__getSelectedPages(),
      total: pageNodes.length,
    }, '*');
  } catch { /* parent gone */ }
};
const paintPicks = () => {
  for (const wrap of scroll.querySelectorAll('.card-wrap')) {
    let anyOn = false;
    let anyOff = false;
    for (const btn of wrap.querySelectorAll('.preview-page-pick')) {
      const n = Number(btn.dataset.page);
      const on = pageIsOn(n);
      btn.setAttribute('aria-pressed', on ? 'true' : 'false');
      if (on) anyOn = true;
      else anyOff = true;
    }
    wrap.classList.toggle('is-off', anyOff && !anyOn);
  }
};
const togglePage = (n) => {
  const total = pageNodes.length;
  if (!(n >= 1 && n <= total)) return;
  if (!selectedPages) selectedPages = new Set(allPageNums(total));
  if (selectedPages.has(n)) selectedPages.delete(n);
  else selectedPages.add(n);
  if (selectedPages.size === total) selectedPages = null;
  paintPicks();
  notifySelection();
};
const makePickBtn = (pageNo) => {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'preview-page-pick';
  btn.dataset.page = String(pageNo);
  btn.textContent = String(pageNo);
  btn.setAttribute('aria-pressed', 'true');
  btn.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    togglePage(pageNo);
  });
  return btn;
};
window.__getSelectedPages = () =>
  selectedPages ? [...selectedPages].sort((a, b) => a - b) : allPageNums(pageNodes.length);
window.__getLogicalPageCount = () => pageNodes.length;
window.__selectAllPages = () => {
  selectedPages = null;
  paintPicks();
  return true;
};
window.__setSelectedPages = (list) => {
  selectedPages = new Set();
  const total = pageNodes.length;
  if (Array.isArray(list)) {
    for (const raw of list) {
      const n = Math.round(Number(raw));
      if (n >= 1 && n <= total) selectedPages.add(n);
    }
  }
  if (selectedPages.size === total) selectedPages = null;
  paintPicks();
  return true;
};
window.__getCurrentPage = () => {
  const cards = [...scroll.querySelectorAll('.card-wrap, :scope > .card')];
  if (cards.length === 0) return 1;
  const mid = scroll.scrollTop + scroll.clientHeight * 0.35;
  let current = 1;
  for (let i = 0; i < cards.length; i++) {
    const top = cards[i].offsetTop;
    const bottom = top + cards[i].offsetHeight;
    if (mid >= top && mid < bottom) return i + 1;
    if (mid >= bottom) current = i + 1;
  }
  return Math.min(cards.length, Math.max(1, current));
};
/// Fit paper cards into the left pane width (default zoom-to-fit).
const paneViewScale = (paperW) => {
  const paneW = Math.max(scroll.clientWidth - 24, 40);
  if (!(paperW > 0)) return 1;
  return Math.max(0.12, Math.min(4, (paneW / paperW) * userZoom));
};
const applyPrintLayoutNow = (layout) => {
  if (!layout || !scroll || pageNodes.length === 0) return false;
  pendingLayout = layout;
  const paperW = Number(layout.pageW) * CSS_PX;
  const paperH = Number(layout.pageH) * CSS_PX;
  const view = paneViewScale(paperW);
  const padT = Number(layout.top) * CSS_PX * view;
  const padR = Number(layout.right) * CSS_PX * view;
  const padB = Number(layout.bottom) * CSS_PX * view;
  const padL = Number(layout.left) * CSS_PX * view;
  const dispW = Math.max(1, paperW * view);
  const dispH = Math.max(1, paperH * view);
  const innerW = Math.max(1, dispW - padL - padR);
  const innerH = Math.max(1, dispH - padT - padB);
  const scale = Math.min(2, Math.max(0.1, Number(layout.scale) || 1));
  const cols = Math.max(1, Math.round(Number(layout.nupCols)) || 1);
  const rows = Math.max(1, Math.round(Number(layout.nupRows)) || 1);
  const perSheet = cols * rows;
  const direction =
    layout.layoutDirection === 'rtl-ttb' ||
    layout.layoutDirection === 'ttb-ltr' ||
    layout.layoutDirection === 'ttb-rtl'
      ? layout.layoutDirection
      : 'ltr-ttb';
  const sheetTransforms = [];
  if (layout.reverseOrientation) sheetTransforms.push('rotate(180deg)');
  if (layout.flipHorizontal) sheetTransforms.push('scaleX(-1)');
  const sheetTransform = sheetTransforms.length ? sheetTransforms.join(' ') : '';
  const hairline = layout.pageBorder === 'hairline';
  scroll.replaceChildren();
  for (let i = 0; i < pageNodes.length; i += perSheet) {
    const card = document.createElement('div');
    card.className = 'card';
    card.style.width = Math.round(dispW) + 'px';
    card.style.height = Math.round(dispH) + 'px';
    card.style.padding =
      Math.round(padT) + 'px ' + Math.round(padR) + 'px ' +
      Math.round(padB) + 'px ' + Math.round(padL) + 'px';
    card.style.alignItems = 'stretch';
    card.style.justifyContent = 'stretch';
    if (hairline) {
      card.style.outline = '1px solid #000';
      card.style.outlineOffset = '-1px';
    } else {
      card.style.outline = '';
      card.style.outlineOffset = '';
    }
    const grid = document.createElement('div');
    grid.className = 'nup-grid';
    grid.style.gridTemplateColumns = 'repeat(' + cols + ', 1fr)';
    grid.style.gridTemplateRows = 'repeat(' + rows + ', 1fr)';
    grid.style.transform = sheetTransform;
    grid.style.transformOrigin = 'center center';
    const cells = [];
    for (let c = 0; c < perSheet; c++) {
      const cell = document.createElement('div');
      cell.className = 'nup-cell';
      cells.push(cell);
      grid.appendChild(cell);
    }
    const cellW = innerW / cols;
    const cellH = innerH / rows;
    const chunk = pageNodes.slice(i, i + perSheet);
    for (let p = 0; p < chunk.length; p++) {
      const cell = cells[nupCellIndex(p, cols, rows, direction)];
      if (!cell) continue;
      const node = chunk[p];
      const natW = node.pageWIn * CSS_PX;
      const natH = node.pageHIn * CSS_PX;
      const fit = Math.min(cellW / natW, cellH / natH) * scale;
      const tw = Math.max(1, Math.floor(natW * fit));
      const th = Math.max(1, Math.floor(natH * fit));
      node.canvas.style.width = tw + 'px';
      node.canvas.style.height = th + 'px';
      node.stage.style.width = tw + 'px';
      node.stage.style.height = th + 'px';
      node.stage.style.transform = '';
      cell.appendChild(node.stage);
    }
    card.appendChild(grid);
    const wrap = document.createElement('div');
    wrap.className = 'card-wrap';
    wrap.appendChild(card);
    const picks = document.createElement('div');
    picks.className = 'preview-page-picks';
    for (let p = 0; p < chunk.length; p++) {
      picks.appendChild(makePickBtn(chunk[p].pageIndex));
    }
    wrap.appendChild(picks);
    scroll.appendChild(wrap);
  }
  paintPicks();
  notifySelection();
  return true;
};
window.__applyPrintLayout = (layout) => {
  pendingLayout = layout;
  return applyPrintLayoutNow(layout);
};
window.__fitPreview = () => {
  if (pendingLayout) applyPrintLayoutNow(pendingLayout);
};
const scheduleFit = () => {
  if (fitRaf) return;
  fitRaf = requestAnimationFrame(() => {
    fitRaf = 0;
    window.__fitPreview();
  });
};
const waitForPane = () => new Promise((resolve) => {
  if (scroll.clientWidth >= 40 && scroll.clientHeight >= 40) {
    resolve();
    return;
  }
  const ro = new ResizeObserver(() => {
    if (scroll.clientWidth >= 40 && scroll.clientHeight >= 40) {
      ro.disconnect();
      resolve();
    }
  });
  ro.observe(scroll);
  // Fallback if ResizeObserver never fires (should not happen after setBounds).
  setTimeout(() => { ro.disconnect(); resolve(); }, 2000);
});
try {
  await waitForPane();
  const doc = await pdfjs.getDocument({ url: './preview.pdf', useWasm: false }).promise;
  const total = doc.numPages;
  scroll.replaceChildren();
  // Rasterize once; Layout pane (N-up / flip / border) rearranges via
  // __applyPrintLayout without re-rendering PDF pages. Unblock the parent
  // after the first page so the dialog is not stuck waiting on long jobs.
  const renderScale = 1.5;
  for (let i = 1; i <= total; i++) {
    const page = await doc.getPage(i);
    const base = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: renderScale });
    const stage = document.createElement('div');
    stage.className = 'stage';
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const pageWIn = base.width / 72;
    const pageHIn = base.height / 72;
    canvas.dataset.pageWIn = String(pageWIn);
    canvas.dataset.pageHIn = String(pageHIn);
    canvas.style.width = Math.floor(pageWIn * CSS_PX) + 'px';
    canvas.style.height = Math.floor(pageHIn * CSS_PX) + 'px';
    stage.appendChild(canvas);
    pageNodes.push({ stage, canvas, pageWIn, pageHIn, pageIndex: i });
    const ctx = canvas.getContext('2d');
    await page.render({ canvas, canvasContext: ctx, viewport }).promise;
    page.cleanup();
    showPendingOrFallback();
    if (i === 1) markReady();
  }
  new ResizeObserver(scheduleFit).observe(scroll);
  window.addEventListener('resize', scheduleFit);
  window.addEventListener('wheel', (event) => {
    if (!event.ctrlKey) return;
    event.preventDefault();
    userZoom = Math.max(0.25, Math.min(4, userZoom * Math.exp(-event.deltaY * 0.01)));
    scheduleFit();
  }, { passive: false });
} catch (err) {
  const p = document.createElement('p');
  p.className = 'err';
  p.textContent = String(err && err.message ? err.message : err);
  scroll.appendChild(p);
} finally {
  showPendingOrFallback();
  markReady();
}
</script>
</body></html>`,
    'utf8',
  )
  return htmlPath
}

/// Subset 1-based pages into a new PDF, keeping the requested order
/// (needed for Paper Handling reverse / odd-even).
async function subsetPdfPages(pdf: Buffer, pageList: readonly number[]): Promise<Buffer> {
  const { PDFDocument } = await import('pdf-lib')
  const src = await PDFDocument.load(pdf)
  const count = src.getPageCount()
  const indices = pdfPageIndicesInOrder(pageList, count)
  if (indices === null) {
    throw new Error(`none of the requested pages exist in this ${count}-page document`)
  }
  if (indices === undefined) return pdf
  const out = await PDFDocument.create()
  const copied = await out.copyPages(src, indices)
  for (const page of copied) out.addPage(page)
  return Buffer.from(await out.save())
}

/// Lay out as PDF, subset pages, then silent-print with system-dialog options
/// (printer / copies / duplex / color) — no second system print panel.
async function printFromPrintDialog(args: {
  htmlPath: string
  workDir: string
  choice: PrintDialogChoice
  sourceRequest: WorkbookExportPdfRequest | undefined
  renderPdf: PrintPdfRenderer | undefined
}): Promise<WorkbookPrintResult> {
  const pdf = await layoutPdfBytes({
    htmlPath: args.htmlPath,
    workDir: args.workDir,
    fileName: 'print-job.html',
    choice: args.choice,
    sourceRequest: args.sourceRequest,
    renderPdf: args.renderPdf,
  })
  return printPdfBytesSilent(pdf, printOptionsFromDialog(args.choice), args.workDir)
}

async function printPdfBytesSilent(
  pdf: Buffer,
  options: WebContentsPrintOptions,
  workDir: string,
): Promise<WorkbookPrintResult> {
  const pdfPath = join(workDir, 'print-job.pdf')
  await writeFile(pdfPath, pdf)
  const win = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true },
  })
  try {
    await win.loadFile(pdfPath)
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 400)
    })
    const outcome = await new Promise<{ success: boolean; failureReason: string }>((resolve) => {
      win.webContents.print(options, (success, failureReason) =>
        resolve({ success, failureReason }),
      )
    })
    if (outcome.success) return { ok: true }
    return /cancel/i.test(outcome.failureReason)
      ? { ok: false }
      : { ok: false, error: outcome.failureReason }
  } finally {
    if (!win.isDestroyed()) win.destroy()
  }
}

async function previewPageInfo(
  win: BrowserWindow,
): Promise<{ pageCount: number; currentPage: number }> {
  if (win.isDestroyed()) return { pageCount: 1, currentPage: 1 }
  try {
    const info = (await win.webContents.executeJavaScript(`(() => {
      const frame = document.getElementById('preview-frame')
      const doc = frame && frame.contentDocument
      if (!doc) return { pageCount: 1, currentPage: 1 }
      const w = frame && frame.contentWindow
      const logical =
        w && typeof w.__getLogicalPageCount === 'function' ? w.__getLogicalPageCount() : 0
      const current =
        w && typeof w.__getCurrentPage === 'function' ? w.__getCurrentPage() : 1
      const slots = [...doc.querySelectorAll('#preview-pages .preview-card-slot')]
      const cards = slots.length
        ? slots
        : [...doc.querySelectorAll('#scroll .card')]
      const total = Math.max(1, logical || cards.length)
      return { pageCount: total, currentPage: Math.max(1, current) }
    })()`)) as { pageCount: number; currentPage: number }
    if (
      info &&
      Number.isFinite(info.pageCount) &&
      Number.isFinite(info.currentPage)
    ) {
      return {
        pageCount: Math.max(1, Math.round(info.pageCount)),
        currentPage: Math.max(1, Math.round(info.currentPage)),
      }
    }
  } catch {
    /* preview torn down */
  }
  return { pageCount: 1, currentPage: 1 }
}

type FreezeCandidate = {
  index: number
  x: number
  y: number
  width: number
  height: number
}

/// Capture live visible preview pages into bitmaps. Only pages already in view
/// (no forced scrolling), so flinging through large page counts stays smooth.
async function freezeVisiblePreviewPages(
  win: BrowserWindow,
  generation: { current: number },
  mine: number,
): Promise<void> {
  if (win.isDestroyed() || generation.current !== mine) return
  let candidates: FreezeCandidate[] = []
  try {
    candidates = (await win.webContents.executeJavaScript(`(() => {
      const frame = document.getElementById('preview-frame')
      if (!frame || !frame.contentWindow) return []
      const w = frame.contentWindow
      if (typeof w.__freezeCandidates !== 'function') return []
      const host = frame.getBoundingClientRect()
      const list = w.__freezeCandidates() || []
      return list.map((c) => ({
        index: c.index,
        x: c.x + host.x,
        y: c.y + host.y,
        width: c.width,
        height: c.height,
      }))
    })()`)) as FreezeCandidate[]
  } catch {
    return
  }
  if (!Array.isArray(candidates) || candidates.length === 0) return
  for (const candidate of candidates) {
    if (win.isDestroyed() || generation.current !== mine) return
    if (
      !candidate ||
      !Number.isFinite(candidate.index) ||
      candidate.width < 8 ||
      candidate.height < 8
    ) {
      continue
    }
    // One frame for the live card to paint before capture.
    await new Promise<void>((resolve) => setTimeout(resolve, 16))
    if (win.isDestroyed() || generation.current !== mine) return
    try {
      const image = await win.webContents.capturePage({
        x: Math.max(0, Math.round(candidate.x)),
        y: Math.max(0, Math.round(candidate.y)),
        width: Math.min(Math.round(candidate.width), 4096),
        height: Math.min(Math.round(candidate.height), 4096),
      })
      if (image.isEmpty() || generation.current !== mine) continue
      const dataUrl = image.toDataURL()
      await win.webContents.executeJavaScript(
        `(() => {
          const w = document.getElementById('preview-frame')?.contentWindow
          return w && w.__applyPageBitmap && w.__applyPageBitmap(${Math.floor(candidate.index)}, ${JSON.stringify(dataUrl)})
        })()`,
      )
    } catch {
      /* resize / navigation race — skip this page */
    }
  }
}

/// Sheets print HTML + preview-fit script so the iframe exposes __applyPrintLayout.
async function writeWorkbookPreviewHtml(workDir: string, sourceHtmlPath: string): Promise<string> {
  const source = await readFile(sourceHtmlPath, 'utf8')
  const outPath = join(workDir, 'preview-frame.html')
  const script = `<script>${PREVIEW_FIT_SCRIPT}</script>`
  const html = source.includes('</body>')
    ? source.replace('</body>', `${script}</body>`)
    : `${source}${script}`
  await writeFile(outPath, html, 'utf8')
  return outPath
}

export type PrintDialogMode = 'workbook' | 'document'

export async function openCombinedPrintDialog(args: {
  parent: BrowserWindow | null
  workDir: string
  settings: PrintDialogDefaults
  /// Suggested PDF file name when the destination is Save as PDF.
  fileName: string
  /// Workbook HTML preview (required when mode is `workbook`).
  htmlPath?: string
  /// Pre-built PDF preview/job (required when mode is `document`, e.g. Word).
  pdfPath?: string
  mode?: PrintDialogMode
  printContent?: PrintContentMeta
  rebuildContent?: PrintContentRebuild
  sourceRequest?: WorkbookExportPdfRequest
  renderPdf?: PrintPdfRenderer
}): Promise<WorkbookPrintResult> {
  const documentMode = args.mode === 'document' || Boolean(args.pdfPath && !args.htmlPath)
  if (documentMode && !args.pdfPath) {
    return { ok: false, error: label('appPrintFailed') }
  }
  if (!documentMode && !args.htmlPath) {
    return { ok: false, error: label('appPrintFailed') }
  }
  const channel = `sheets:print-dialog:${randomUUID()}`
  const layoutChannel = `sheets:print-dialog-layout:${randomUUID()}`
  const contentChannel = `sheets:print-dialog-content:${randomUUID()}`
  const preloadPath = join(args.workDir, 'print-dialog-preload.js')
  const panelPath = join(args.workDir, 'print-dialog.html')
  await writeFile(preloadPath, preloadSource(channel, layoutChannel, contentChannel), 'utf8')
  const printContentMeta: PrintContentMeta = args.printContent ?? {
    scope: 'active',
    sheetCount: 1,
    hasSelection: false,
  }
  /// Coalesce omitted center flags to the app default (both on).
  let liveSettings: PrintDialogDefaults = {
    ...args.settings,
    horizontalCentered: args.settings.horizontalCentered ?? true,
    verticalCentered: args.settings.verticalCentered ?? true,
  }
  let liveRequest = args.sourceRequest
  let documentPageCount = 1

  const parent = args.parent && !args.parent.isDestroyed() ? args.parent : undefined
  const win = new BrowserWindow({
    width: DIALOG_WIDTH,
    height: 760,
    minWidth: DIALOG_MIN_WIDTH,
    minHeight: 520,
    show: false,
    title: label('appPrintDialogTitle'),
    autoHideMenuBar: true,
    ...(parent ? { parent, modal: true } : {}),
    webPreferences: {
      /// Off: sandboxed windows cannot read iframe contentWindow.
      sandbox: false,
      contextIsolation: true,
      nodeIntegration: false,
      /// Off: Chromium treats each file:// URL as a unique origin, so the
      /// sidebar could not call `__applyPrintLayout` inside the preview iframe.
      webSecurity: false,
      preload: preloadPath,
    },
  })

  let layoutGen = 0
  let fittedPage = ''
  let previewFramePath = ''
  const freezeGen = { current: 0 }
  let freezeBusy = false
  const freezeTimer = documentMode
    ? undefined
    : setInterval(() => {
        if (freezeBusy || win.isDestroyed()) return
        const mine = freezeGen.current
        freezeBusy = true
        void freezeVisiblePreviewPages(win, freezeGen, mine).finally(() => {
          freezeBusy = false
        })
      }, 120)
  /// Apply layout into the panel's preview iframe (same origin as the sidebar).
  const applyPreviewLayout = async (layout: PreviewLayout): Promise<boolean> => {
    if (win.isDestroyed()) return false
    const patch = previewLayoutPatch(layout)
    const pageKey = `${patch.pageW}x${patch.pageH}`
    const gen = ++layoutGen
    freezeGen.current += 1
    const patchJson = JSON.stringify(patch)
    try {
      const ok = (await win.webContents.executeJavaScript(
        `(async () => {
          const frame = document.getElementById('preview-frame')
          if (!frame) return false
          const deadline = Date.now() + 8000
          while (Date.now() < deadline) {
            const w = frame.contentWindow
            if (w && typeof w.__applyPrintLayout === 'function') {
              // First page ready is enough — do not wait for every page to rasterize.
              if (w.__pdfReady) {
                await Promise.race([
                  w.__pdfReady,
                  new Promise((r) => setTimeout(r, Math.max(0, deadline - Date.now()))),
                ])
              }
              w.__applyPrintLayout(${patchJson})
              if (typeof w.__fitPreview === 'function') w.__fitPreview()
              return true
            }
            await new Promise((r) => setTimeout(r, 32))
          }
          return false
        })()`,
      )) === true
      if (gen !== layoutGen || win.isDestroyed()) return ok
      if (ok) fittedPage = pageKey
      return ok
    } catch {
      return false
    }
  }

  /// sheet/width modes use the preview-measured scale so print matches the
  /// left pane; scale mode keeps the user’s percent.
  const resolvePrintScale = async (choice: PrintDialogChoice): Promise<number> => {
    if (choice.fitMode === 'scale') return choice.scale
    if (win.isDestroyed()) return choice.scale
    try {
      const scale = (await win.webContents.executeJavaScript(
        `(() => {
          const w = document.getElementById('preview-frame')?.contentWindow
          return w && typeof w.__contentScale === 'number' ? w.__contentScale : null
        })()`,
      )) as number | null
      if (typeof scale === 'number' && Number.isFinite(scale) && scale >= 0.1 && scale <= 2) {
        return Math.round(scale * 1000) / 1000
      }
    } catch {
      /* preview torn down */
    }
    return choice.scale
  }

  const reloadPreviewFrame = async (absolutePath: string): Promise<void> => {
    previewFramePath = absolutePath
    if (win.isDestroyed()) return
    const name = basename(absolutePath)
    await win.webContents.executeJavaScript(
      `(() => {
        const frame = document.getElementById('preview-frame')
        if (!frame) return
        frame.src = ${JSON.stringify(name + '?t=' + Date.now())}
      })()`,
    )
  }

  let settled = false
  const choice = await new Promise<PrintDialogChoice | { error: string } | null>((resolve) => {
    let names = new Set<string>()
    let onChoice: (event: IpcMainEvent, payload: unknown) => void = () => {}
    const finish = (value: PrintDialogChoice | { error: string } | null) => {
      if (settled) return
      settled = true
      freezeGen.current += 1
      if (freezeTimer !== undefined) clearInterval(freezeTimer)
      ipcMain.removeListener(channel, onChoice)
      ipcMain.removeHandler(layoutChannel)
      ipcMain.removeHandler(contentChannel)
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
    ipcMain.on(channel, onChoice)
    ipcMain.handle(layoutChannel, async (_event, payload: unknown) => {
      const next = parsePreviewLayout(payload)
      if (!next) return { ok: false, error: 'invalid layout' }
      const ok = await applyPreviewLayout(next)
      return { ok }
    })
    ipcMain.handle(contentChannel, async (_event, scopeRaw: unknown) => {
      if (documentMode) return { error: label('appPrintFailed') }
      if (
        scopeRaw !== 'active' &&
        scopeRaw !== 'workbook' &&
        scopeRaw !== 'selection'
      ) {
        return { error: label('appPrintFailed') }
      }
      if (!args.rebuildContent || !args.htmlPath) return { error: label('appPrintFailed') }
      try {
        const rebuilt = await args.rebuildContent(scopeRaw)
        if ('error' in rebuilt) return { error: rebuilt.error }
        liveSettings = rebuilt.settings
        liveRequest = rebuilt.request
        printContentMeta.scope = scopeRaw
        const nextPreview = await writeWorkbookPreviewHtml(args.workDir, args.htmlPath)
        await reloadPreviewFrame(nextPreview)
        await applyPreviewLayout(liveSettings)
        const { pageCount } = await previewPageInfo(win)
        return { ok: true, pageCount }
      } catch (error: unknown) {
        return { error: error instanceof Error ? error.message : label('appPrintFailed') }
      }
    })
    win.once('closed', () => finish(null))
    void (async () => {
      try {
        /// Prefer the parent window's printer list — a blank dialog
        /// has stalled getPrintersAsync on some macOS setups.
        const printerSource =
          parent && !parent.isDestroyed() ? parent.webContents : win.webContents
        const listed = await printerSource.getPrintersAsync()
        const printers: ListedPrinter[] = listed.map(listedPrinterOf)
        names = new Set(printers.map((printer) => printer.name))
        let pageCount = 1
        let previewSrc = 'preview-frame.html'
        if (documentMode && args.pdfPath) {
          documentPageCount = await readPdfPageCount(args.pdfPath)
          pageCount = documentPageCount
          previewFramePath = await writePdfJsPreviewHtml(args.workDir, args.pdfPath)
          previewSrc = basename(previewFramePath)
        } else if (args.htmlPath) {
          previewFramePath = await writeWorkbookPreviewHtml(args.workDir, args.htmlPath)
          previewSrc = basename(previewFramePath)
          const raw = await readFile(args.htmlPath, 'utf8')
          pageCount = Math.max(1, (raw.match(/class="area"/g) ?? []).length)
        }
        await writeFile(
          panelPath,
          panelHtml(
            printers,
            liveSettings,
            pageCount,
            printContentMeta,
            documentMode,
            previewSrc,
          ),
          'utf8',
        )
        await win.loadFile(panelPath)
        /// Show immediately — document PDF previews rasterize in the iframe and
        /// must not block the dialog behind an invisible window.
        if (!win.isDestroyed()) {
          win.show()
          win.focus()
        }
        await applyPreviewLayout(liveSettings)
        if (!documentMode && !win.isDestroyed()) {
          pageCount = (await previewPageInfo(win)).pageCount
          await win.webContents.executeJavaScript(
            `(() => {
              const to = document.getElementById('page-to')
              const from = document.getElementById('page-from')
              if (to) { to.max = String(${pageCount}); if (!to.value) to.value = String(${pageCount}) }
              if (from) from.max = String(${pageCount})
            })()`,
          )
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
    const pageInfo = await previewPageInfo(win)
    const pageCount = documentMode
      ? Math.max(1, documentPageCount || pageInfo.pageCount)
      : pageInfo.pageCount
    const { currentPage } = pageInfo
    const ranged = resolvePrintPageList(choice, pageCount, currentPage)
    if (ranged === null) {
      if (!win.isDestroyed()) win.destroy()
      return { ok: false, error: label('appPrintFailed') }
    }
    const pageList = applyPaperHandlingPages(
      ranged,
      pageCount,
      choice.sheetsToPrint,
      choice.pageOrder,
    )
    if (pageList && pageList.length === 0) {
      if (!win.isDestroyed()) win.destroy()
      return { ok: false, error: label('appPrintFailed') }
    }
    const scaledChoice: PrintDialogChoice = {
      ...choice,
      scale: documentMode ? choice.scale : await resolvePrintScale(choice),
      ...(pageList ? { pageList } : {}),
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
      if (documentMode && args.pdfPath) {
        let pdf: Buffer = await readFile(args.pdfPath)
        if (scaledChoice.pageList) {
          pdf = Buffer.from(await subsetPdfPages(pdf, scaledChoice.pageList))
        }
        await atomicWriteFile(selection.filePath, pdf)
        return { ok: true, path: selection.filePath }
      }
      if (!args.htmlPath) return { ok: false, error: label('appPrintFailed') }
      return await savePdfFromPrintDialog({
        htmlPath: args.htmlPath,
        workDir: args.workDir,
        outPath: selection.filePath,
        choice: scaledChoice,
        sourceRequest: liveRequest,
        renderPdf: args.renderPdf,
      })
    }

    if (!win.isDestroyed()) await closePrintWindow(win)
    if (documentMode && args.pdfPath) {
      let pdf: Buffer = await readFile(args.pdfPath)
      if (scaledChoice.pageList) {
        pdf = Buffer.from(await subsetPdfPages(pdf, scaledChoice.pageList))
      }
      return await printPdfBytesSilent(
        pdf,
        printOptionsFromDialog(scaledChoice),
        args.workDir,
      )
    }
    if (!args.htmlPath) return { ok: false, error: label('appPrintFailed') }
    return await printFromPrintDialog({
      htmlPath: args.htmlPath,
      workDir: args.workDir,
      choice: scaledChoice,
      sourceRequest: liveRequest,
      renderPdf: args.renderPdf,
    })
  } catch (error: unknown) {
    if (!win.isDestroyed()) win.destroy()
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

/// Word / document print: same combined dialog as Sheets, left pane loads a
/// pre-built PDF (pagination already done by the editor).
export async function openDocumentPrintDialog(args: {
  parent: BrowserWindow | null
  pdfPath: string
  workDir: string
  fileName: string
  settings?: PrintDialogDefaults
}): Promise<WorkbookPrintResult> {
  return openCombinedPrintDialog({
    parent: args.parent,
    workDir: args.workDir,
    fileName: args.fileName,
    pdfPath: args.pdfPath,
    mode: 'document',
    settings: args.settings ?? {
      landscape: false,
      pageSize: 'A4',
      margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
      scale: 1,
      fitMode: 'scale',
      horizontalCentered: true,
      verticalCentered: true,
    },
  })
}

/// Markdown / HTML / PDF / Word: write bytes into a temp folder, open the
/// combined print dialog, then delete the folder.
export async function openDocumentPrintDialogFromPdf(args: {
  parent: BrowserWindow | null
  pdf: Buffer
  fileName: string
}): Promise<WorkbookPrintResult> {
  const workDir = await mkdtemp(join(tmpdir(), 'genoffice-print-dlg-'))
  const base = args.fileName.replace(/\.pdf$/i, '').trim() || 'Document'
  try {
    const pdfPath = join(workDir, 'preview.pdf')
    await writeFile(pdfPath, args.pdf)
    return await openDocumentPrintDialog({
      parent: args.parent,
      pdfPath,
      workDir,
      fileName: `${base}.pdf`,
    })
  } finally {
    await rm(workDir, { recursive: true, force: true }).catch(() => undefined)
  }
}
