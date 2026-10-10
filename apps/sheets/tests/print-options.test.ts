import { describe, expect, it } from 'vitest'
import {
  SAVE_AS_PDF_DEVICE,
  applyPaperHandlingPages,
  parsePrintDialogChoice,
  pdfPageIndicesInOrder,
  printOptionsFor,
  printOptionsFromDialog,
  printToPdfOptionsFromDialog,
  requestFromPrintDialog,
  resolvePrintPageList,
  scaleToFitFactor,
} from '../src/main/print-options'
import { previewLayoutCss, previewLayoutPatch } from '../src/shared/preview-layout'
import type { WorkbookExportPdfRequest } from '../src/shared/desktop-api'

const REQUEST: WorkbookExportPdfRequest = {
  fileName: 'Book1.pdf',
  html: '<html></html>',
  landscape: true,
  pageSize: 'A4',
  margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
  scale: 0.8,
}

describe('printOptionsFor', () => {
  it('carries the page setup over in the units webContents.print expects', () => {
    expect(printOptionsFor(REQUEST)).toEqual({
      silent: false,
      printBackground: true,
      landscape: true,
      pageSize: 'A4',
      margins: { marginType: 'custom', top: 72, bottom: 72, left: 67, right: 67 },
      scaleFactor: 80,
    })
  })

  it('converts a custom paper size from inches to microns', () => {
    expect(printOptionsFor({ ...REQUEST, pageSize: { width: 8.5, height: 14 } }).pageSize).toEqual({
      width: 215900,
      height: 355600,
    })
  })
})

const CHOICE = {
  deviceName: 'HP',
  copies: 2,
  collate: true,
  color: false,
  duplexMode: 'longEdge' as const,
  rangeMode: 'all' as const,
  customRange: '',
  landscape: false,
  pageSize: 'Letter' as const,
  margins: { top: 1, bottom: 1, left: 0.5, right: 0.5 },
  scale: 1,
  fitMode: 'sheet' as const,
  horizontalCentered: true,
  verticalCentered: false,
  pagesPerSheet: 1,
  layoutDirection: 'ltr-ttb' as const,
  pageBorder: 'none' as const,
  reverseOrientation: false,
  flipHorizontal: false,
  sheetsToPrint: 'all' as const,
  pageOrder: 'automatic' as const,
  scaleToFit: false,
  destinationPageSize: 'Letter' as const,
  scaleDownOnly: false,
}

describe('parsePrintDialogChoice', () => {
  it('keeps a known printer and system-dialog options', () => {
    expect(parsePrintDialogChoice(CHOICE, new Set(['HP']))).toEqual(CHOICE)
  })

  it('rejects a printer that is not in the list', () => {
    expect(parsePrintDialogChoice(CHOICE, new Set(['Other']))).toBeNull()
  })

  it('accepts Save as PDF without a printer list entry', () => {
    expect(
      parsePrintDialogChoice({ ...CHOICE, deviceName: SAVE_AS_PDF_DEVICE }, new Set()),
    ).toMatchObject({ deviceName: SAVE_AS_PDF_DEVICE })
  })
})

describe('resolvePrintPageList', () => {
  it('prints all pages when mode is all', () => {
    expect(resolvePrintPageList({ rangeMode: 'all', customRange: '' }, 5, 2)).toBeUndefined()
  })

  it('picks the current preview page', () => {
    expect(resolvePrintPageList({ rangeMode: 'current', customRange: '' }, 5, 3)).toEqual([3])
  })

  it('parses a custom range into 1-based pages', () => {
    expect(resolvePrintPageList({ rangeMode: 'custom', customRange: '1,3,5-6' }, 10, 1)).toEqual([
      1, 3, 5, 6,
    ])
  })

  it('rejects an invalid custom range', () => {
    expect(resolvePrintPageList({ rangeMode: 'custom', customRange: '99' }, 5, 1)).toBeNull()
  })
})

describe('previewLayoutCss', () => {
  it('centers on the page in print and keeps the paper card on screen', () => {
    const css = previewLayoutCss({
      landscape: false,
      pageSize: 'A4',
      margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
      scale: 0.8,
      fitMode: 'sheet',
      horizontalCentered: true,
      verticalCentered: true,
    })
    expect(css).toContain('--page-w: 8.27')
    expect(css).toContain('--page-h: 11.69')
    expect(css).toContain('--m-top: 0.75')
    expect(css).toContain('--m-left: 0.7')
    expect(css).toContain('--print-scale: 0.8')
    expect(css).toContain('--align-x: center')
    expect(css).toContain('--align-y: center')
  })

  it('turns alignment off without changing the paper size', () => {
    expect(
      previewLayoutPatch({
        landscape: true,
        pageSize: 'A4',
        margins: { top: 1, bottom: 0.5, left: 0.25, right: 0 },
        scale: 1,
        fitMode: 'scale',
        horizontalCentered: false,
        verticalCentered: false,
      }),
    ).toEqual({
      pageW: 11.69,
      pageH: 8.27,
      top: 1,
      right: 0,
      bottom: 0.5,
      left: 0.25,
      scale: 1,
      fitMode: 'scale',
      align: 'flex-start',
      justify: 'flex-start',
    })
  })
})

describe('printOptionsFromDialog', () => {
  it('silent-prints with system-dialog destination options and no extra margins', () => {
    expect(printOptionsFromDialog(CHOICE)).toEqual({
      silent: true,
      printBackground: true,
      deviceName: 'HP',
      copies: 2,
      collate: true,
      color: false,
      duplexMode: 'longEdge',
      pagesPerSheet: 1,
      margins: { marginType: 'none' },
    })
  })
})

describe('printToPdfOptionsFromDialog', () => {
  it('uses printToPDF units; ranges are applied by subsetting later', () => {
    expect(printToPdfOptionsFromDialog(CHOICE)).toEqual({
      landscape: false,
      pageSize: 'Letter',
      margins: { top: 1, bottom: 1, left: 0.5, right: 0.5 },
      scale: 1,
      printBackground: true,
    })
  })

  it('switches to destination paper and scales when scale-to-fit is on', () => {
    expect(
      printToPdfOptionsFromDialog({
        ...CHOICE,
        pageSize: 'Letter',
        destinationPageSize: 'A4',
        scaleToFit: true,
        scaleDownOnly: true,
        scale: 1,
      }),
    ).toMatchObject({
      pageSize: 'A4',
      scale: scaleToFitFactor('Letter', 'A4', false, true),
    })
  })
})

describe('applyPaperHandlingPages', () => {
  it('keeps full document order as undefined', () => {
    expect(applyPaperHandlingPages(undefined, 4, 'all', 'automatic')).toBeUndefined()
  })

  it('filters odd pages from a range', () => {
    expect(applyPaperHandlingPages([1, 2, 3, 4], 4, 'odd', 'normal')).toEqual([1, 3])
  })

  it('reverses page order', () => {
    expect(applyPaperHandlingPages(undefined, 3, 'all', 'reverse')).toEqual([3, 2, 1])
  })

  it('returns an empty list when no even pages exist', () => {
    expect(applyPaperHandlingPages([1], 1, 'even', 'normal')).toEqual([])
  })
})

describe('pdfPageIndicesInOrder', () => {
  it('skips a rewrite when pages are already in document order', () => {
    expect(pdfPageIndicesInOrder([1, 2, 3], 3)).toBeUndefined()
  })

  it('keeps reverse order instead of sorting', () => {
    expect(pdfPageIndicesInOrder([3, 2, 1], 3)).toEqual([2, 1, 0])
  })

  it('rejects a list with no in-range pages', () => {
    expect(pdfPageIndicesInOrder([], 3)).toBeNull()
  })
})

describe('requestFromPrintDialog', () => {
  it('keeps header templates while applying destination paper', () => {
    const request: WorkbookExportPdfRequest = {
      ...REQUEST,
      headerTemplate: '<div>H</div>',
    }
    expect(
      requestFromPrintDialog(request, {
        ...CHOICE,
        scaleToFit: true,
        destinationPageSize: 'A4',
      }),
    ).toMatchObject({
      headerTemplate: '<div>H</div>',
      pageSize: 'A4',
      landscape: false,
    })
  })
})
