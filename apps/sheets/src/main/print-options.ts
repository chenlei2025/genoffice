import type { PrintToPDFOptions, WebContentsPrintOptions } from 'electron'
import type { WorkbookExportPdfRequest } from '../shared/desktop-api'

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

/// Page setup the dialog opens with: the sheet's effective print settings.
export interface PrintDialogDefaults {
  landscape: boolean
  pageSize: DialogPageSize
  margins: PrintDialogMargins
  /// 0.1–2, same unit as the export request.
  scale: number
  horizontalCentered: boolean
  verticalCentered: boolean
}

export interface PrintDialogChoice extends PrintDialogDefaults {
  deviceName: string
  copies: number
  collate: boolean
  color: boolean
  duplexMode: 'simplex' | 'shortEdge' | 'longEdge'
  /// 1-based inclusive range. Absent prints every page.
  pages?: { from: number; to: number }
}

export interface PreviewLayoutInput {
  landscape: boolean
  pageSize: DialogPageSize
  margins: PrintDialogMargins
  scale: number
  horizontalCentered: boolean
  verticalCentered: boolean
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

function pagesOf(value: unknown): { from: number; to: number } | undefined | null {
  if (value == null) return undefined
  if (typeof value !== 'object') return null
  const raw = value as { from?: unknown; to?: unknown }
  const from = Math.round(Number(raw.from))
  const to = Math.round(Number(raw.to))
  if (!Number.isInteger(from) || !Number.isInteger(to)) return null
  if (from < 1 || to < from || to > 9999) return null
  return { from, to }
}

export function isSaveAsPdfDevice(deviceName: string): boolean {
  return deviceName === SAVE_AS_PDF_DEVICE
}

/// Accepts the sidebar payload. Unknown printers and unusable page setup
/// are rejected so the job cannot be sent to a device the user did not pick.
/// `SAVE_AS_PDF_DEVICE` is always accepted (save dialog + printToPDF).
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
  const pages = pagesOf(raw.pages)
  const scale = Number(raw.scale)
  if (!pageSize || !margins || pages === null) return null
  if (!Number.isFinite(scale) || scale < 0.1 || scale > 2) return null
  if (
    typeof raw.landscape !== 'boolean' ||
    typeof raw.color !== 'boolean' ||
    typeof raw.collate !== 'boolean' ||
    typeof raw.horizontalCentered !== 'boolean' ||
    typeof raw.verticalCentered !== 'boolean'
  ) {
    return null
  }
  return {
    deviceName: raw.deviceName,
    copies: Math.min(99, Math.max(1, copies)),
    collate: raw.collate,
    color: raw.color,
    duplexMode: raw.duplexMode,
    landscape: raw.landscape,
    pageSize,
    margins,
    scale,
    horizontalCentered: raw.horizontalCentered,
    verticalCentered: raw.verticalCentered,
    ...(pages ? { pages } : {}),
  }
}

/// Live preview update: paper, margins, scale, and on-page alignment. Rejects
/// a payload that is not a usable page setup.
export function parsePreviewLayout(payload: unknown): PreviewLayoutInput | null {
  if (!payload || typeof payload !== 'object') return null
  const raw = payload as Record<string, unknown>
  const pageSize = pageSizeOf(raw.pageSize)
  const margins = marginsOf(raw.margins)
  const scale = clampScale(raw.scale)
  if (!pageSize || !margins || scale === null || typeof raw.landscape !== 'boolean') return null
  if (typeof raw.horizontalCentered !== 'boolean' || typeof raw.verticalCentered !== 'boolean') {
    return null
  }
  return {
    landscape: raw.landscape,
    pageSize,
    margins,
    scale,
    horizontalCentered: raw.horizontalCentered,
    verticalCentered: raw.verticalCentered,
  }
}

/// Silent print options for the chosen printer, including the page setup the
/// user confirmed in the dialog. `omitPageSetup` drops paper size and margins
/// for a retry when the printer rejects a custom size.
export function printOptionsFromDialog(
  choice: PrintDialogChoice,
  omitPageSetup = false,
): WebContentsPrintOptions {
  const laidOut = printOptionsFor({
    fileName: 'print.pdf',
    html: '<html></html>',
    landscape: choice.landscape,
    pageSize: choice.pageSize,
    margins: choice.margins,
    scale: choice.scale,
  })
  const { pageSize: _pageSize, margins: _margins, ...withoutPaper } = laidOut
  return {
    ...(omitPageSetup ? withoutPaper : laidOut),
    silent: true,
    deviceName: choice.deviceName,
    copies: choice.copies,
    collate: choice.collate,
    color: choice.color,
    duplexMode: choice.duplexMode,
    ...(choice.pages
      ? { pageRanges: [{ from: choice.pages.from - 1, to: choice.pages.to - 1 }] }
      : {}),
  }
}

/// printToPDF options for the dialog's Save as PDF destination. Same page
/// setup units as File → Export PDF (inches / named paper / 0.1–2 scale).
export function printToPdfOptionsFromDialog(choice: PrintDialogChoice): PrintToPDFOptions {
  return {
    landscape: choice.landscape,
    pageSize: choice.pageSize,
    margins: choice.margins,
    scale: choice.scale,
    printBackground: true,
    ...(choice.pages ? { pageRanges: `${choice.pages.from}-${choice.pages.to}` } : {}),
  }
}
