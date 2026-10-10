import type { PrintToPDFOptions, WebContentsPrintOptions } from 'electron'
import type { WorkbookExportPdfRequest } from '../shared/desktop-api'
import { parsePrintRange } from '../renderer/print-range'

const MICRONS_PER_INCH = 25_400
const CSS_PX_PER_INCH = 96

/// Sentinel device name for the dialog's "Save as PDF" destination. Not a
/// CUPS printer — parsePrintDialogChoice accepts it without a printer list hit.
export const SAVE_AS_PDF_DEVICE = '__save-as-pdf__'

export const NAMED_PAGE_SIZES = ['A3', 'A4', 'A5', 'Legal', 'Letter', 'Tabloid'] as const

export type NamedPageSize = (typeof NAMED_PAGE_SIZES)[number]

export type DialogPageSize = NamedPageSize | { width: number; height: number }

export interface PrintDialogMargins {
  top: number
  bottom: number
  left: number
  right: number
}

/// How content is scaled onto the paper (from sheet page setup).
export type PrintFitMode = 'scale' | 'sheet' | 'width'

/// Page setup the dialog opens with: the sheet's effective print settings.
export interface PrintDialogDefaults {
  landscape: boolean
  pageSize: DialogPageSize
  margins: PrintDialogMargins
  /// 0.1–2; used when fitMode is `scale`.
  scale: number
  /// `sheet` = whole sheet on one page; `width` = all columns on one page.
  fitMode: PrintFitMode
  horizontalCentered: boolean
  verticalCentered: boolean
}

export type PrintRangeMode = 'all' | 'current' | 'custom'

/// N-up page order on each sheet (WPS / system Layout pane).
export type PrintLayoutDirection = 'ltr-ttb' | 'rtl-ttb' | 'ttb-ltr' | 'ttb-rtl'

export type PrintPageBorder = 'none' | 'hairline'

/// Paper Handling “待打印张数”.
export type PrintSheetsToPrint = 'all' | 'odd' | 'even'

/// Paper Handling “纸张顺序”.
export type PrintPageOrder = 'automatic' | 'normal' | 'reverse'

/// Sidebar choice: system-print-dialog controls + sheet + layout + paper handling.
export interface PrintDialogChoice extends PrintDialogDefaults {
  deviceName: string
  copies: number
  collate: boolean
  color: boolean
  duplexMode: 'simplex' | 'shortEdge' | 'longEdge'
  rangeMode: PrintRangeMode
  /// Raw custom range text when rangeMode is `custom` (e.g. "1,3,5-8").
  customRange: string
  /// 1-based pages resolved at confirm time; absent prints every page.
  pageList?: number[]
  /// Layout pane (system print “Layout”).
  pagesPerSheet: number
  layoutDirection: PrintLayoutDirection
  pageBorder: PrintPageBorder
  reverseOrientation: boolean
  flipHorizontal: boolean
  /// Paper Handling pane.
  sheetsToPrint: PrintSheetsToPrint
  pageOrder: PrintPageOrder
  scaleToFit: boolean
  /// Used when scaleToFit is on (destination media).
  destinationPageSize: DialogPageSize
  scaleDownOnly: boolean
}

export interface PreviewLayoutInput {
  landscape: boolean
  pageSize: DialogPageSize
  margins: PrintDialogMargins
  scale: number
  fitMode: PrintFitMode
  horizontalCentered: boolean
  verticalCentered: boolean
}

function fitModeOf(value: unknown): PrintFitMode | null {
  return value === 'scale' || value === 'sheet' || value === 'width' ? value : null
}

const NAMED = new Set<string>(NAMED_PAGE_SIZES)

/// webContents.print options matching what printToPDF gets for the same
/// request: print() wants custom sizes in microns, margins in CSS pixels and
/// the scale as a percentage.
export function printOptionsFor(request: WorkbookExportPdfRequest): WebContentsPrintOptions {
  const { margins, pageSize } = request
  return {
    silent: false,
    printBackground: true,
    landscape: request.landscape,
    pageSize:
      typeof pageSize === 'string'
        ? pageSize
        : {
            width: Math.round(pageSize.width * MICRONS_PER_INCH),
            height: Math.round(pageSize.height * MICRONS_PER_INCH),
          },
    margins: {
      marginType: 'custom',
      top: Math.round(margins.top * CSS_PX_PER_INCH),
      bottom: Math.round(margins.bottom * CSS_PX_PER_INCH),
      left: Math.round(margins.left * CSS_PX_PER_INCH),
      right: Math.round(margins.right * CSS_PX_PER_INCH),
    },
    scaleFactor: Math.round(request.scale * 100),
  }
}

function clampInches(value: unknown): number | null {
  const n = Number(value)
  if (!Number.isFinite(n)) return null
  return Math.min(3, Math.max(0, Math.round(n * 100) / 100))
}

function clampScale(value: unknown): number | null {
  const n = Number(value)
  if (!Number.isFinite(n)) return null
  return Math.min(2, Math.max(0.1, Math.round(n * 1000) / 1000))
}

function pageSizeOf(value: unknown): DialogPageSize | null {
  if (typeof value === 'string') return NAMED.has(value) ? (value as NamedPageSize) : null
  if (!value || typeof value !== 'object') return null
  const size = value as { width?: unknown; height?: unknown }
  const width = Number(size.width)
  const height = Number(size.height)
  if (!Number.isFinite(width) || !Number.isFinite(height)) return null
  if (width <= 0 || height <= 0 || width > 100 || height > 100) return null
  return { width, height }
}

function marginsOf(value: unknown): PrintDialogMargins | null {
  if (!value || typeof value !== 'object') return null
  const raw = value as Record<string, unknown>
  const top = clampInches(raw.top)
  const bottom = clampInches(raw.bottom)
  const left = clampInches(raw.left)
  const right = clampInches(raw.right)
  if (top === null || bottom === null || left === null || right === null) return null
  return { top, bottom, left, right }
}

function rangeModeOf(value: unknown): PrintRangeMode | null {
  return value === 'all' || value === 'current' || value === 'custom' ? value : null
}

function layoutDirectionOf(value: unknown): PrintLayoutDirection | null {
  return value === 'ltr-ttb' ||
    value === 'rtl-ttb' ||
    value === 'ttb-ltr' ||
    value === 'ttb-rtl'
    ? value
    : null
}

function pageBorderOf(value: unknown): PrintPageBorder | null {
  return value === 'none' || value === 'hairline' ? value : null
}

function sheetsToPrintOf(value: unknown): PrintSheetsToPrint | null {
  return value === 'all' || value === 'odd' || value === 'even' ? value : null
}

function pageOrderOf(value: unknown): PrintPageOrder | null {
  return value === 'automatic' || value === 'normal' || value === 'reverse' ? value : null
}

export function isSaveAsPdfDevice(deviceName: string): boolean {
  return deviceName === SAVE_AS_PDF_DEVICE
}

/// Resolve page range into 1-based page numbers. `null` = invalid custom input;
/// `undefined` = print every page.
export function resolvePrintPageList(
  choice: Pick<PrintDialogChoice, 'rangeMode' | 'customRange'>,
  pageCount: number,
  currentPage: number,
): number[] | undefined | null {
  if (!(pageCount >= 1)) return undefined
  if (choice.rangeMode === 'all') return undefined
  if (choice.rangeMode === 'current') {
    const page = Math.min(pageCount, Math.max(1, Math.round(currentPage)))
    return [page]
  }
  const zeroBased = parsePrintRange(choice.customRange.trim(), pageCount)
  if (!zeroBased) return null
  return zeroBased.map((i) => i + 1)
}

/// Apply Paper Handling odd/even + reverse order to a 1-based page list.
/// `undefined` input means every page in document order.
export function applyPaperHandlingPages(
  pages: number[] | undefined,
  pageCount: number,
  sheetsToPrint: PrintSheetsToPrint,
  pageOrder: PrintPageOrder,
): number[] | undefined {
  if (!(pageCount >= 1)) return pages
  let list = pages ?? Array.from({ length: pageCount }, (_x, i) => i + 1)
  if (sheetsToPrint === 'odd') list = list.filter((p) => p % 2 === 1)
  else if (sheetsToPrint === 'even') list = list.filter((p) => p % 2 === 0)
  if (pageOrder === 'reverse') list = [...list].reverse()
  if (list.length === 0) return []
  if (
    sheetsToPrint === 'all' &&
    pageOrder !== 'reverse' &&
    list.length === pageCount &&
    list.every((p, i) => p === i + 1)
  ) {
    return undefined
  }
  return list
}

/// 0-based indices in the caller’s order. `null` = nothing printable;
/// `undefined` = every page already in document order (skip a rewrite).
export function pdfPageIndicesInOrder(
  pageList: readonly number[],
  pageCount: number,
): number[] | undefined | null {
  const indices: number[] = []
  const seen = new Set<number>()
  for (const n of pageList) {
    if (!Number.isInteger(n) || n < 1 || n > pageCount) continue
    const index = n - 1
    if (seen.has(index)) continue
    seen.add(index)
    indices.push(index)
  }
  if (indices.length === 0) return null
  if (indices.length === pageCount && indices.every((index, i) => index === i)) return undefined
  return indices
}

/// Accepts the system-dialog-style sidebar payload.
export function parsePrintDialogChoice(
  payload: unknown,
  printerNames: ReadonlySet<string>,
): PrintDialogChoice | null {
  if (!payload || typeof payload !== 'object') return null
  const raw = payload as Record<string, unknown>
  if (typeof raw.deviceName !== 'string') return null
  if (!isSaveAsPdfDevice(raw.deviceName) && !printerNames.has(raw.deviceName)) return null
  const copies = Math.round(Number(raw.copies))
  if (!Number.isFinite(copies)) return null
  if (raw.duplexMode !== 'simplex' && raw.duplexMode !== 'shortEdge' && raw.duplexMode !== 'longEdge') {
    return null
  }
  const pageSize = pageSizeOf(raw.pageSize)
  const margins = marginsOf(raw.margins)
  const scale = Number(raw.scale)
  const fitMode = fitModeOf(raw.fitMode)
  const rangeMode = rangeModeOf(raw.rangeMode)
  if (!pageSize || !margins || !fitMode || !rangeMode) return null
  if (!Number.isFinite(scale) || scale < 0.1 || scale > 2) return null
  const pagesPerSheet = Math.round(Number(raw.pagesPerSheet))
  const layoutDirection = layoutDirectionOf(raw.layoutDirection)
  const pageBorder = pageBorderOf(raw.pageBorder)
  const sheetsToPrint = sheetsToPrintOf(raw.sheetsToPrint)
  const pageOrder = pageOrderOf(raw.pageOrder)
  const destinationPageSize = pageSizeOf(raw.destinationPageSize) ?? pageSize
  if (
    typeof raw.landscape !== 'boolean' ||
    typeof raw.color !== 'boolean' ||
    typeof raw.collate !== 'boolean' ||
    typeof raw.horizontalCentered !== 'boolean' ||
    typeof raw.verticalCentered !== 'boolean' ||
    typeof raw.reverseOrientation !== 'boolean' ||
    typeof raw.flipHorizontal !== 'boolean' ||
    typeof raw.scaleToFit !== 'boolean' ||
    typeof raw.scaleDownOnly !== 'boolean' ||
    !layoutDirection ||
    !pageBorder ||
    !sheetsToPrint ||
    !pageOrder ||
    !destinationPageSize ||
    !Number.isFinite(pagesPerSheet)
  ) {
    return null
  }
  const customRange = typeof raw.customRange === 'string' ? raw.customRange : ''
  return {
    deviceName: raw.deviceName,
    copies: Math.min(99, Math.max(1, copies)),
    collate: raw.collate,
    color: raw.color,
    duplexMode: raw.duplexMode,
    rangeMode,
    customRange,
    landscape: raw.landscape,
    pageSize,
    margins,
    scale,
    fitMode,
    horizontalCentered: raw.horizontalCentered,
    verticalCentered: raw.verticalCentered,
    pagesPerSheet: Math.min(16, Math.max(1, pagesPerSheet)),
    layoutDirection,
    pageBorder,
    reverseOrientation: raw.reverseOrientation,
    flipHorizontal: raw.flipHorizontal,
    sheetsToPrint,
    pageOrder,
    scaleToFit: raw.scaleToFit,
    destinationPageSize,
    scaleDownOnly: raw.scaleDownOnly,
  }
}

/// Live preview update: paper, margins, scale, and on-page alignment.
export function parsePreviewLayout(payload: unknown): PreviewLayoutInput | null {
  if (!payload || typeof payload !== 'object') return null
  const raw = payload as Record<string, unknown>
  const pageSize = pageSizeOf(raw.pageSize)
  const margins = marginsOf(raw.margins)
  const scale = clampScale(raw.scale)
  const fitMode = fitModeOf(raw.fitMode)
  if (!pageSize || !margins || scale === null || !fitMode || typeof raw.landscape !== 'boolean') {
    return null
  }
  if (typeof raw.horizontalCentered !== 'boolean' || typeof raw.verticalCentered !== 'boolean') {
    return null
  }
  return {
    landscape: raw.landscape,
    pageSize,
    margins,
    scale,
    fitMode,
    horizontalCentered: raw.horizontalCentered,
    verticalCentered: raw.verticalCentered,
  }
}

/// Silent print of an already-laid-out PDF: system-dialog destination options
/// only (page geometry is baked into the PDF bytes).
/// `margins: none` is required — default/silent FitToPrintableArea would
/// re-margin the PDF and clip left/right content that the preview showed.
export function printOptionsFromDialog(choice: PrintDialogChoice): WebContentsPrintOptions {
  return {
    silent: true,
    printBackground: true,
    deviceName: choice.deviceName,
    copies: choice.copies,
    collate: choice.collate,
    color: choice.color,
    duplexMode: choice.duplexMode,
    pagesPerSheet: choice.pagesPerSheet,
    margins: { marginType: 'none' },
  }
}

const PAPER_INCHES: Record<NamedPageSize, { width: number; height: number }> = {
  Letter: { width: 8.5, height: 11 },
  Tabloid: { width: 11, height: 17 },
  Legal: { width: 8.5, height: 14 },
  A3: { width: 11.69, height: 16.54 },
  A4: { width: 8.27, height: 11.69 },
  A5: { width: 5.83, height: 8.27 },
}

function pageInches(size: DialogPageSize): { width: number; height: number } {
  return typeof size === 'string' ? PAPER_INCHES[size] : size
}

/// Scale factor from source paper onto destination paper (Paper Handling).
export function scaleToFitFactor(
  source: DialogPageSize,
  destination: DialogPageSize,
  landscape: boolean,
  scaleDownOnly: boolean,
): number {
  const src = pageInches(source)
  const dst = pageInches(destination)
  const srcW = landscape ? src.height : src.width
  const srcH = landscape ? src.width : src.height
  const dstW = landscape ? dst.height : dst.width
  const dstH = landscape ? dst.width : dst.height
  if (!(srcW > 0) || !(srcH > 0) || !(dstW > 0) || !(dstH > 0)) return 1
  let factor = Math.min(dstW / srcW, dstH / srcH)
  if (scaleDownOnly && factor > 1) factor = 1
  return Math.round(factor * 1000) / 1000
}

/// printToPDF options for layout / Save as PDF. Ranges are applied by subsetting.
export function printToPdfOptionsFromDialog(choice: PrintDialogChoice): PrintToPDFOptions {
  const pageSize = choice.scaleToFit ? choice.destinationPageSize : choice.pageSize
  let scale = choice.scale
  if (choice.scaleToFit) {
    scale =
      Math.round(
        scale *
          scaleToFitFactor(
            choice.pageSize,
            choice.destinationPageSize,
            choice.landscape,
            choice.scaleDownOnly,
          ) *
          1000,
      ) / 1000
  }
  return {
    landscape: choice.landscape,
    pageSize,
    margins: choice.margins,
    scale,
    printBackground: true,
  }
}

/// Overlay dialog paper/scale onto the laid-out export request (keeps headers).
export function requestFromPrintDialog(
  request: WorkbookExportPdfRequest,
  choice: PrintDialogChoice,
): WorkbookExportPdfRequest {
  const pdf = printToPdfOptionsFromDialog(choice)
  return {
    ...request,
    landscape: choice.landscape,
    pageSize: choice.scaleToFit ? choice.destinationPageSize : choice.pageSize,
    margins: choice.margins,
    scale: typeof pdf.scale === 'number' ? pdf.scale : choice.scale,
    horizontalCentered: choice.horizontalCentered,
    verticalCentered: choice.verticalCentered,
  }
}
