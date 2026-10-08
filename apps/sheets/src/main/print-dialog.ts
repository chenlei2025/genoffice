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

const SIDEBAR_PX = 300

interface ListedPrinter {
  name: string
  displayName: string
}

type DialogLabelKey =
  | 'appPrintDialogTitle'
  | 'appPrintPrinter'
  | 'appPrintSaveAsPdf'
  | 'appPrintCopies'
  | 'appPrintNow'
  | 'appPrintNoPrinter'
  | 'appPrintFailed'
  | 'appExportPdfTitle'
  | 'appCancel'
  | 'appPrintProperties'
  | 'appPrintPages'
  | 'appPrintAllPages'
  | 'appPrintPageFrom'
  | 'appPrintPageTo'
  | 'appPrintScale'
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
  | 'appOrientationLabel'
  | 'appPortrait'
  | 'appLandscape'
  | 'appSizeLabel'
  | 'appPrintMarginsInches'
  | 'appPrintAlign'
  | 'appPrintCenterHorizontal'
  | 'appPrintCenterVertical'
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
  const named = NAMED_PAGE_SIZES.map(
    (name) =>
      `<option value="${name}"${defaults.pageSize === name ? ' selected' : ''}>${name}</option>`,
  ).join('')
  const customOption = `<option value="custom"${custom ? ' selected' : ''}>${escapeHtml(label('appPrintCustomPaper'))}</option>`
  return named + customOption
}

function panelHtml(printers: readonly ListedPrinter[], defaults: PrintDialogDefaults): string {
  const data = JSON.stringify({ printers, defaults }).replace(/</g, '\\u003c')
  const none = printers.length === 0
  const custom = typeof defaults.pageSize === 'object' ? defaults.pageSize : { width: 8.5, height: 11 }
  const scalePercent = Math.round(defaults.scale * 100)
  const inch = (value: number) => String(Math.round(value * 100) / 100)
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
    gap: 10px;
    padding: 16px;
  }
  h2 { margin: 4px 0 0; font-size: 12px; font-weight: 600; }
  label { display: flex; flex-direction: column; gap: 4px; }
  label.check { flex-direction: row; align-items: center; gap: 8px; }
  select, input { font: inherit; padding: 6px 8px; }
  .pair, .margins { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .hint { margin: 0; opacity: 0.8; }
  .actions {
    flex: none;
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    justify-content: flex-end;
    padding: 12px 16px;
    border-top: 1px solid color-mix(in srgb, CanvasText 18%, transparent);
    background: Canvas;
  }
  button { font: inherit; padding: 6px 14px; }
  button.primary { background: AccentColor; color: AccentColorText; border: 0; }
  button:disabled { opacity: 0.5; }
</style>
</head>
<body>
<aside>
<div class="fields">
  <label>${escapeHtml(label('appPrintPrinter'))}
    <select id="printer"></select>
  </label>
  <div class="pair" id="printer-copies">
    <label>${escapeHtml(label('appPrintCopies'))}
      <input id="copies" type="number" min="1" max="99" value="1">
    </label>
    <label class="check" style="align-self: end; padding-bottom: 6px;">
      <input id="collate" type="checkbox" checked>
      ${escapeHtml(label('appPrintCollate'))}
    </label>
  </div>
  <label>${escapeHtml(label('appPrintPages'))}
    <select id="pages">
      <option value="all">${escapeHtml(label('appPrintAllPages'))}</option>
      <option value="range">${escapeHtml(label('appPrintPageFrom'))}–${escapeHtml(label('appPrintPageTo'))}</option>
    </select>
  </label>
  <div class="pair" id="range" hidden>
    <label>${escapeHtml(label('appPrintPageFrom'))}
      <input id="page-from" type="number" min="1" max="9999" value="1">
    </label>
    <label>${escapeHtml(label('appPrintPageTo'))}
      <input id="page-to" type="number" min="1" max="9999" value="1">
    </label>
  </div>
  <h2>${escapeHtml(label('appPrintProperties'))}</h2>
  <label>${escapeHtml(label('appSizeLabel'))}
    <select id="paper">${paperOptions(defaults)}</select>
  </label>
  <div class="pair" id="custom-paper"${typeof defaults.pageSize === 'string' ? ' hidden' : ''}>
    <label>${escapeHtml(label('appWidth'))}
      <input id="paper-width" type="number" min="1" max="100" step="0.01" value="${custom.width}">
    </label>
    <label>${escapeHtml(label('appHeight'))}
      <input id="paper-height" type="number" min="1" max="100" step="0.01" value="${custom.height}">
    </label>
  </div>
  <label>${escapeHtml(label('appOrientationLabel'))}
    <select id="orientation">
      <option value="portrait"${defaults.landscape ? '' : ' selected'}>${escapeHtml(label('appPortrait'))}</option>
      <option value="landscape"${defaults.landscape ? ' selected' : ''}>${escapeHtml(label('appLandscape'))}</option>
    </select>
  </label>
  <label>${escapeHtml(label('appPrintScale'))}
    <input id="scale" type="number" min="10" max="200" step="1" value="${scalePercent}">
  </label>
  <div>
    <div>${escapeHtml(label('appPrintMarginsInches'))}</div>
    <div class="margins">
      <label>${escapeHtml(label('appPrintMarginTop'))}
        <input id="margin-top" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.top)}">
      </label>
      <label>${escapeHtml(label('appPrintMarginBottom'))}
        <input id="margin-bottom" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.bottom)}">
      </label>
      <label>${escapeHtml(label('appPrintMarginLeft'))}
        <input id="margin-left" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.left)}">
      </label>
      <label>${escapeHtml(label('appPrintMarginRight'))}
        <input id="margin-right" type="number" min="0" max="3" step="0.05" value="${inch(defaults.margins.right)}">
      </label>
    </div>
  </div>
  <div>
    <div>${escapeHtml(label('appPrintAlign'))}</div>
    <label class="check">
      <input id="center-h" type="checkbox"${defaults.horizontalCentered ? ' checked' : ''}>
      ${escapeHtml(label('appPrintCenterHorizontal'))}
    </label>
    <label class="check">
      <input id="center-v" type="checkbox"${defaults.verticalCentered ? ' checked' : ''}>
      ${escapeHtml(label('appPrintCenterVertical'))}
    </label>
  </div>
  <label id="color-field">${escapeHtml(label('appPrintColorMode'))}
    <select id="color">
      <option value="1">${escapeHtml(label('appPrintColor'))}</option>
      <option value="0">${escapeHtml(label('appPrintGrayscale'))}</option>
    </select>
  </label>
  <label id="duplex-field">${escapeHtml(label('appPrintSides'))}
    <select id="duplex">
      <option value="simplex">${escapeHtml(label('appPrintOneSided'))}</option>
      <option value="longEdge">${escapeHtml(label('appPrintLongEdge'))}</option>
      <option value="shortEdge">${escapeHtml(label('appPrintShortEdge'))}</option>
    </select>
  </label>
  <p class="hint" id="no-printer"${none ? '' : ' hidden'}>${escapeHtml(label('appPrintNoPrinter'))}</p>
</div>
  <div class="actions">
    <button id="cancel" type="button">${escapeHtml(label('appCancel'))}</button>
    <button id="save-pdf" type="button">${escapeHtml(label('appPrintSaveAsPdf'))}</button>
    <button id="print" class="primary" type="button">${escapeHtml(label('appPrintNow'))}</button>
  </div>
</aside>
<script id="dialog-data" type="application/json">${data}</script>
<script>
  const data = JSON.parse(document.getElementById('dialog-data').textContent)
  const SAVE_AS_PDF = ${JSON.stringify(SAVE_AS_PDF_DEVICE)}
  const select = document.getElementById('printer')
  const pdfOption = document.createElement('option')
  pdfOption.value = SAVE_AS_PDF
  pdfOption.textContent = ${JSON.stringify(label('appPrintSaveAsPdf'))}
  select.appendChild(pdfOption)
  for (const printer of data.printers) {
    const option = document.createElement('option')
    option.value = printer.name
    option.textContent = printer.displayName || printer.name
    select.appendChild(option)
  }
  if (data.printers.length > 0) select.value = data.printers[0].name
  const pages = document.getElementById('pages')
  const range = document.getElementById('range')
  pages.onchange = () => { range.hidden = pages.value !== 'range' }
  const paper = document.getElementById('paper')
  const customPaper = document.getElementById('custom-paper')
  paper.onchange = () => { customPaper.hidden = paper.value !== 'custom' }
  const printerOnly = [
    document.getElementById('printer-copies'),
    document.getElementById('color-field'),
    document.getElementById('duplex-field'),
  ]
  const noPrinter = document.getElementById('no-printer')
  const printBtn = document.getElementById('print')
  const syncDestination = () => {
    const pdf = select.value === SAVE_AS_PDF
    for (const el of printerOnly) el.hidden = pdf
    noPrinter.hidden = pdf || data.printers.length > 0
    printBtn.disabled = !pdf && data.printers.length === 0
  }
  select.onchange = syncDestination
  syncDestination()
  const num = (id) => Number(document.getElementById(id).value)
  const currentLayout = () => ({
    landscape: document.getElementById('orientation').value === 'landscape',
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
    horizontalCentered: document.getElementById('center-h').checked,
    verticalCentered: document.getElementById('center-v').checked,
  })
  // input fires while spinning or typing so the preview tracks margins,
  // scale, and alignment immediately; change covers selects and checkboxes.
  const publishLayout = () => window.printDialog.layout(currentLayout())
  for (const id of ['paper', 'orientation', 'scale', 'center-h', 'center-v', 'paper-width', 'paper-height', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right']) {
    const el = document.getElementById(id)
    el.addEventListener('input', publishLayout)
    el.addEventListener('change', publishLayout)
  }
  const submitChoice = (deviceName) => {
    const pageSize = paper.value === 'custom'
      ? { width: num('paper-width'), height: num('paper-height') }
      : paper.value
    let pageRange = null
    if (pages.value === 'range') {
      let from = Math.round(num('page-from'))
      let to = Math.round(num('page-to'))
      if (from > to) { const swap = from; from = to; to = swap }
      pageRange = { from, to }
    }
    window.printDialog.submit({
      deviceName,
      copies: Math.round(num('copies')) || 1,
      collate: document.getElementById('collate').checked,
      color: document.getElementById('color').value === '1',
      duplexMode: document.getElementById('duplex').value,
      landscape: document.getElementById('orientation').value === 'landscape',
      pageSize,
      margins: {
        top: num('margin-top'),
        bottom: num('margin-bottom'),
        left: num('margin-left'),
        right: num('margin-right'),
      },
      scale: (Math.round(num('scale')) || 100) / 100,
      pages: pageRange,
      horizontalCentered: document.getElementById('center-h').checked,
      verticalCentered: document.getElementById('center-v').checked,
    })
  }
  document.getElementById('cancel').onclick = () => window.printDialog.submit(null)
  document.getElementById('save-pdf').onclick = () => submitChoice(SAVE_AS_PDF)
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

/// Fits the paper sheet in the preview and lets a trackpad pinch change that
/// scale. Zoom lives on `#preview-pages`, which print media forces back to 1.
const PREVIEW_FIT_SCRIPT = `(() => {
  if (window.__fitPreview) return
  let userZoom = 1
  const fitHost = document.createElement('div')
  fitHost.id = 'preview-fit'
  const pages = document.createElement('div')
  pages.id = 'preview-pages'
  const areas = [...document.querySelectorAll('.area')]
  for (const area of areas) pages.appendChild(area)
  fitHost.appendChild(pages)
  document.body.appendChild(fitHost)
  const fit = () => {
    const page = pages.querySelector('.area')
    if (!page) return
    const previous = pages.style.zoom
    pages.style.zoom = '1'
    const width = page.offsetWidth
    const height = page.offsetHeight
    if (fitHost.clientWidth < 40 || fitHost.clientHeight < 40 || !width || !height) {
      if (previous) pages.style.zoom = previous
      return
    }
    const fitted = Math.min((fitHost.clientWidth - 32) / width, (fitHost.clientHeight - 32) / height)
    pages.style.zoom = String(Math.max(0.2, Math.min(4, fitted * userZoom)))
  }
  window.__fitPreview = fit
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
    'landscape' | 'pageSize' | 'margins' | 'scale' | 'horizontalCentered' | 'verticalCentered'
  >,
): Promise<void> {
  const layoutCss = previewLayoutCss({
    landscape: choice.landscape,
    pageSize: choice.pageSize,
    margins: choice.margins,
    scale: choice.scale,
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

/// Save-as-PDF from the print dialog. Uses a hidden window (same path as
/// File → Export PDF) so the preview's zoom / modal chrome cannot empty the
/// printable area or block the save panel.
async function savePdfFromPrintDialog(args: {
  parent: BrowserWindow | null | undefined
  htmlPath: string
  workDir: string
  fileName: string
  choice: PrintDialogChoice
}): Promise<WorkbookPrintResult> {
  const parent = args.parent && !args.parent.isDestroyed() ? args.parent : undefined
  const selection = await showSaveDialogWithMemory(dialog, parent, {
    title: label('appExportPdfTitle'),
    defaultPath: args.fileName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  })
  if (selection.canceled || !selection.filePath) return { ok: false }

  const pdfHtmlPath = join(args.workDir, 'print-save.pdf.html')
  await writeJobHtml(args.htmlPath, pdfHtmlPath, args.choice)
  const hidden = await openJobWindow(pdfHtmlPath)
  try {
    const pdf = await hidden.webContents.printToPDF(printToPdfOptionsFromDialog(args.choice))
    await atomicWriteFile(selection.filePath, pdf)
    return { ok: true, path: selection.filePath }
  } finally {
    if (!hidden.isDestroyed()) hidden.destroy()
  }
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
            sheet.style.zoom = String(layout.scale)
          }
        })()`,
      )
      .then(() => {
        if (gen !== layoutGen || preview.webContents.isDestroyed()) return
        if (pageKey === fittedPage) return
        fittedPage = pageKey
        refitPreview()
      })
      .catch(() => undefined)
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

  if (!choice || 'error' in choice || win.isDestroyed()) {
    if (!win.isDestroyed()) win.destroy()
    return choice && 'error' in choice ? { ok: false, error: choice.error } : { ok: false }
  }

  // Close the modal preview first so save/print are not blocked, and so the
  // job runs from clean HTML (same layout CSS the left pane was simulating).
  if (!win.isDestroyed()) win.destroy()
  try {
    if (isSaveAsPdfDevice(choice.deviceName)) {
      return await savePdfFromPrintDialog({
        parent,
        htmlPath: args.htmlPath,
        workDir: args.workDir,
        fileName: args.fileName,
        choice,
      })
    }
    return await printFromPrintDialog({
      htmlPath: args.htmlPath,
      workDir: args.workDir,
      choice,
    })
  } catch (error: unknown) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}
