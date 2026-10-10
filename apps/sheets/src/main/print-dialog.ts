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
/// orientation toggles, PDF + Cancel + Print footer). Left preview unchanged.
function panelHtml(
  printers: readonly ListedPrinter[],
  defaults: PrintDialogDefaults,
  pageCount: number,
  printContent: PrintContentMeta,
  documentMode = false,
): string {
  const data = JSON.stringify({
    printers,
    defaults,
    pageCount,
    printContent,
    documentMode,
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
    background: Canvas;
    color: CanvasText;
  }
  [hidden] { display: none !important; }
  aside {
    position: fixed;
    top: 0; right: 0; bottom: 0;
    width: ${SIDEBAR_PX}px;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    min-height: 0;
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
  rangeAll.onchange = syncRange
  rangeCurrent.onchange = syncRange
  rangeFromTo.onchange = syncRange
  pageFrom.onfocus = () => { rangeFromTo.checked = true; syncRange() }
  pageTo.onfocus = () => { rangeFromTo.checked = true; syncRange() }
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
  for (const head of document.querySelectorAll('.section.collapsible > .section-head')) {
    head.addEventListener('click', () => {
      const section = head.closest('.section')
      const collapsed = section.classList.toggle('collapsed')
      head.setAttribute('aria-expanded', collapsed ? 'false' : 'true')
    })
  }
  // Keep wheel/trackpad scrolling inside the sidebar so the left preview
  // WebContentsView never receives scroll-chained events from this pane.
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
  const currentLayout = () => {
    const d = data.defaults
    const centerH = document.getElementById('center-h')
    const centerV = document.getElementById('center-v')
    return {
      landscape: document.getElementById('orient-landscape').checked,
      pageSize: scaleToFit.checked ? destPaper.value : sourcePageSize(),
      margins: d.margins,
      scale: (Math.round(num('scale')) || Math.round(d.scale * 100) || 100) / 100,
      fitMode: currentFitMode(),
      horizontalCentered: centerH ? centerH.checked : !!d.horizontalCentered,
      verticalCentered: centerV ? centerV.checked : !!d.verticalCentered,
    }
  }
  const publishLayout = () => {
    if (data.documentMode) return
    if (window.printDialog && window.printDialog.layout) window.printDialog.layout(currentLayout())
  }
  for (const id of [
    'paper', 'orient-portrait', 'orient-landscape', 'paper-width', 'paper-height',
    'scale', 'scale-mode-pct', 'scale-mode-fit', 'fit-adjust', 'center-h', 'center-v',
    'dest-paper',
  ]) {
    const el = document.getElementById(id)
    if (!el) continue
    el.addEventListener('input', publishLayout)
    el.addEventListener('change', publishLayout)
  }
  const submitChoice = (deviceName) => {
    if (!window.printDialog || !deviceName) return
    printBtn.disabled = true
    savePdfBtn.disabled = true
    document.getElementById('cancel').disabled = true
    const d = data.defaults
    const layout = currentLayout()
    window.printDialog.submit({
      deviceName,
      copies: Math.round(num('copies')) || 1,
      collate: document.getElementById('collate').checked,
      color: document.getElementById('color').value === '1',
      duplexMode: document.getElementById('duplex').value,
      rangeMode: rangeMode(),
      customRange: customRangeValue(),
      landscape: layout.landscape,
      pageSize: sourcePageSize(),
      margins: d.margins,
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
</script>
</body>
</html>`
}

function preloadSource(channel: string, layoutChannel: string, contentChannel: string): string {
  return `const { contextBridge, ipcRenderer } = require('electron')
contextBridge.exposeInMainWorld('printDialog', {
  submit: (choice) => ipcRenderer.send(${JSON.stringify(channel)}, choice),
  layout: (layout) => ipcRenderer.send(${JSON.stringify(layoutChannel)}, layout),
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
  const VIRTUALIZE_AFTER = 1
  const VISIBLE_BUFFER = 0
  const MOUNT_IDLE_MS = 70
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
    cardMetrics = [...pages.children].map((slot) => ({
      top: slot.offsetTop,
      height: slot.offsetHeight,
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

  const mountCard = (spec) => {
    const source = sources[spec.sourceIndex]
    if (!source) return null
    const card = source.cloneNode(true)
    applyPageChrome(card)
    const cardSheet = card.querySelector('.sheet')
    if (!cardSheet) {
      applyAreaPaneTransform(card)
      return card
    }
    applySheetScale(cardSheet, spec.zoom)
    if (spec.pageCount > 1) {
      card.style.justifyContent = 'flex-start'
      card.style.alignItems = spec.alignX
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
    applyAreaPaneTransform(card)
    return card
  }

  const clearBitmaps = () => {
    for (const slot of pages.children) {
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
      const slot = pages.children[i]
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
        pageSpecs.push({
          sourceIndex,
          pageIndex,
          pageCount,
          zoom,
          ah,
          alignX,
        })
        const slot = document.createElement('div')
        slot.className = 'preview-card-slot'
        slot.dataset.empty = '1'
        slot.dataset.live = '0'
        slot.dataset.frozen = '0'
        pages.appendChild(slot)
      }
    }
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
    const needsSize = !pages.firstElementChild?.style.width
    if (!scaleChanged && !needsSize) {
      cacheCardMetrics()
      return true
    }
    if (scaleChanged) clearBitmaps()
    paneScale = next
    const slotW = Math.ceil(pageWpx * paneScale) + 'px'
    const slotH = Math.ceil(pageHpx * paneScale) + 'px'
    for (const slot of pages.children) {
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
  /// Live+mostly-visible pages ready for main-process capturePage.
  window.__freezeCandidates = () => {
    if (!built || scrolling) return []
    const host = scroll.getBoundingClientRect()
    const out = []
    for (let i = 0; i < pages.childElementCount; i++) {
      const slot = pages.children[i]
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
  window.__applyPageBitmap = (index, dataUrl) => {
    const slot = pages.children[index]
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
  preview: WebContentsView,
): Promise<{ pageCount: number; currentPage: number }> {
  if (preview.webContents.isDestroyed()) return { pageCount: 1, currentPage: 1 }
  try {
    const info = (await preview.webContents.executeJavaScript(`(() => {
      const slots = [...document.querySelectorAll('#preview-pages > .preview-card-slot')]
      const total = Math.max(1, slots.length)
      const scroll = document.getElementById('preview-scroll')
      let current = 1
      if (scroll && slots.length > 0) {
        const mid = scroll.scrollTop + scroll.clientHeight * 0.35
        for (let i = 0; i < slots.length; i++) {
          const top = slots[i].offsetTop
          const bottom = top + slots[i].offsetHeight
          if (mid >= top && mid < bottom) { current = i + 1; break }
          if (mid >= bottom) current = i + 1
        }
      }
      return { pageCount: total, currentPage: current }
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
  preview: WebContentsView,
  generation: { current: number },
  mine: number,
): Promise<void> {
  if (preview.webContents.isDestroyed() || generation.current !== mine) return
  let candidates: FreezeCandidate[] = []
  try {
    candidates = (await preview.webContents.executeJavaScript(
      'typeof window.__freezeCandidates === "function" ? window.__freezeCandidates() : []',
    )) as FreezeCandidate[]
  } catch {
    return
  }
  if (!Array.isArray(candidates) || candidates.length === 0) return
  for (const candidate of candidates) {
    if (preview.webContents.isDestroyed() || generation.current !== mine) return
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
    if (preview.webContents.isDestroyed() || generation.current !== mine) return
    try {
      const image = await preview.webContents.capturePage({
        x: Math.max(0, Math.round(candidate.x)),
        y: Math.max(0, Math.round(candidate.y)),
        width: Math.min(Math.round(candidate.width), 4096),
        height: Math.min(Math.round(candidate.height), 4096),
      })
      if (image.isEmpty() || generation.current !== mine) continue
      const dataUrl = image.toDataURL()
      await preview.webContents.executeJavaScript(
        `window.__applyPageBitmap && window.__applyPageBitmap(${Math.floor(candidate.index)}, ${JSON.stringify(dataUrl)})`,
      )
    } catch {
      /* resize / navigation race — skip this page */
    }
  }
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
  let liveSettings = args.settings
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
    webPreferences: { sandbox: true, preload: preloadPath },
  })
  /// Document mode loads a PDF into the left pane; Chromium's viewer needs
  /// `plugins: true`. Without it, `loadFile(*.pdf)` can hang forever and the
  /// dialog never reaches `show()` (Word print stuck on "Generating preview…").
  const preview = new WebContentsView({
    webPreferences: { sandbox: true, plugins: documentMode },
  })
  win.contentView.addChildView(preview)

  let refitTimer: ReturnType<typeof setTimeout> | undefined
  const refitPreview = (): void => {
    if (preview.webContents.isDestroyed()) return
    // Pane-scale only (no card rebuild). Debounce: setBounds + resize fire
    // in bursts and each sync JS call janks the UI thread.
    if (refitTimer !== undefined) clearTimeout(refitTimer)
    refitTimer = setTimeout(() => {
      refitTimer = undefined
      if (preview.webContents.isDestroyed()) return
      void preview.webContents
        .executeJavaScript('window.__fitPreview && window.__fitPreview()')
        .catch(() => {})
    }, 48)
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
  const freezeGen = { current: 0 }
  let freezeBusy = false
  const freezeTimer = documentMode
    ? undefined
    : setInterval(() => {
        if (freezeBusy || preview.webContents.isDestroyed()) return
        const mine = freezeGen.current
        freezeBusy = true
        void freezeVisiblePreviewPages(preview, freezeGen, mine).finally(() => {
          freezeBusy = false
        })
      }, 120)
  /// Write margins, scale, and alignment onto the page itself. Inline styles
  /// beat the baked stylesheet so each keystroke moves the sheet inside the
  /// paper card without refitting (which would reset the pinch zoom).
  const applyPreviewLayout = (layout: PreviewLayout): Promise<void> => {
    if (documentMode || preview.webContents.isDestroyed()) return Promise.resolve()
    const patch = previewLayoutPatch(layout)
    const pageKey = `${patch.pageW}x${patch.pageH}`
    const gen = ++layoutGen
    freezeGen.current += 1
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
            sheet.style.zoom = ''
          }
          if (typeof window.__setPrintScale === 'function') {
            window.__setPrintScale(layout.scale, layout.fitMode)
          } else {
            for (const sheet of document.querySelectorAll('.sheet')) {
              sheet.style.transform = 'scale(' + layout.scale + ')'
              sheet.style.transformOrigin = 'top left'
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
      freezeGen.current += 1
      if (freezeTimer !== undefined) clearInterval(freezeTimer)
      ipcMain.removeListener(channel, onChoice)
      ipcMain.removeListener(layoutChannel, onLayout)
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
    onLayout = (_event, payload) => {
      const next = parsePreviewLayout(payload)
      if (!next) return
      void applyPreviewLayout(next)
    }
    ipcMain.on(channel, onChoice)
    ipcMain.on(layoutChannel, onLayout)
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
        await preview.webContents.loadFile(args.htmlPath)
        await preview.webContents.executeJavaScript(PREVIEW_FIT_SCRIPT)
        await applyPreviewLayout(liveSettings)
        layout()
        const { pageCount } = await previewPageInfo(preview)
        return { ok: true, pageCount }
      } catch (error: unknown) {
        return { error: error instanceof Error ? error.message : label('appPrintFailed') }
      }
    })
    win.once('closed', () => finish(null))
    void (async () => {
      try {
        /// Prefer the parent window's printer list — a blank WebContentsView
        /// has stalled getPrintersAsync on some macOS setups.
        const printerSource =
          parent && !parent.isDestroyed() ? parent.webContents : preview.webContents
        const listed = await printerSource.getPrintersAsync()
        const printers: ListedPrinter[] = listed.map(listedPrinterOf)
        names = new Set(printers.map((printer) => printer.name))
        let pageCount = 1
        if (documentMode && args.pdfPath) {
          const { PDFDocument } = await import('pdf-lib')
          const bytes = await readFile(args.pdfPath)
          const pdf = await PDFDocument.load(bytes)
          documentPageCount = Math.max(1, pdf.getPageCount())
          pageCount = documentPageCount
        } else if (args.htmlPath) {
          await preview.webContents.loadFile(args.htmlPath)
          await preview.webContents.executeJavaScript(PREVIEW_FIT_SCRIPT)
          await applyPreviewLayout(liveSettings)
          layout()
          pageCount = (await previewPageInfo(preview)).pageCount
        }
        await writeFile(
          panelPath,
          panelHtml(printers, liveSettings, pageCount, printContentMeta, documentMode),
          'utf8',
        )
        await win.loadFile(panelPath)
        layout()
        if (!win.isDestroyed()) {
          win.show()
          win.focus()
          layout()
        }
        /// Load the PDF after the shell is visible so a stuck Chromium PDF
        /// viewer cannot block the dialog from appearing. `#toolbar=0` hides
        /// the native viewer chrome (download / print / ⋮) — the right sidebar
        /// already owns those actions for Word print.
        if (documentMode && args.pdfPath && !preview.webContents.isDestroyed()) {
          await Promise.race([
            preview.webContents.loadFile(args.pdfPath, { hash: 'toolbar=0' }),
            new Promise<void>((resolve) => setTimeout(resolve, 8000)),
          ])
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
    const pageInfo = documentMode
      ? { pageCount: documentPageCount, currentPage: 1 }
      : await previewPageInfo(preview)
    const { pageCount, currentPage } = pageInfo
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
      horizontalCentered: false,
      verticalCentered: false,
    },
  })
}
