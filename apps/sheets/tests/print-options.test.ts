import { describe, expect, it } from 'vitest'
import {
  SAVE_AS_PDF_DEVICE,
  parsePrintDialogChoice,
  printOptionsFor,
  printOptionsFromDialog,
  printToPdfOptionsFromDialog,
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
  landscape: false,
  pageSize: 'Letter' as const,
  margins: { top: 1, bottom: 1, left: 0.5, right: 0.5 },
  scale: 1,
  horizontalCentered: true,
  verticalCentered: false,
  pages: { from: 2, to: 4 },
}

describe('parsePrintDialogChoice', () => {
  it('keeps a known printer and the page setup', () => {
    expect(parsePrintDialogChoice(CHOICE, new Set(['HP']))).toEqual(CHOICE)
  })

  it('rejects a printer that is not in the list', () => {
    expect(parsePrintDialogChoice(CHOICE, new Set(['Other']))).toBeNull()
  })

  it('drops an empty page range so the whole document prints', () => {
    expect(parsePrintDialogChoice({ ...CHOICE, pages: null }, new Set(['HP']))?.pages).toBeUndefined()
  })

  it('accepts Save as PDF without a printer list entry', () => {
    expect(
      parsePrintDialogChoice({ ...CHOICE, deviceName: SAVE_AS_PDF_DEVICE }, new Set()),
    ).toMatchObject({ deviceName: SAVE_AS_PDF_DEVICE })
  })
})

describe('previewLayoutCss', () => {
  it('centers on the page in print and keeps the paper card on screen', () => {
    const css = previewLayoutCss({
      landscape: false,
      pageSize: 'A4',
      margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
      scale: 0.8,
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
    expect(css).toContain('--v-min: calc((11.69 - 0.75 - 0.75) * 1in)')
    expect(css).toContain('padding: calc(var(--m-top) * 1in)')
    expect(css).toContain('align-items: var(--align-x)')
    expect(css).toContain('justify-content: var(--align-y)')
    expect(css).toContain('zoom: var(--print-scale)')
    expect(css).toContain('flex: none')
    expect(css).toContain('@media screen')
    expect(css).toContain('@media print')
    expect(css).toContain('min-height: var(--v-min)')
    expect(css).toContain('zoom: 1')
  })

  it('turns alignment off without changing the paper size', () => {
    expect(
      previewLayoutPatch({
        landscape: true,
        pageSize: 'A4',
        margins: { top: 1, bottom: 0.5, left: 0.25, right: 0 },
        scale: 1,
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
      align: 'flex-start',
      justify: 'flex-start',
    })
  })
})

describe('printOptionsFromDialog', () => {
  it('sends the confirmed properties in the units webContents.print expects', () => {
    expect(printOptionsFromDialog(CHOICE)).toMatchObject({
      silent: true,
      deviceName: 'HP',
      copies: 2,
      collate: true,
      color: false,
      duplexMode: 'longEdge',
      landscape: false,
      pageSize: 'Letter',
      margins: { marginType: 'custom', top: 96, bottom: 96, left: 48, right: 48 },
      scaleFactor: 100,
      pageRanges: [{ from: 1, to: 3 }],
    })
  })

  it('can retry without a custom paper size', () => {
    const options = printOptionsFromDialog(
      { ...CHOICE, pageSize: { width: 8.5, height: 11 } },
      true,
    )
    expect(options.pageSize).toBeUndefined()
    expect(options.margins).toBeUndefined()
    expect(options.landscape).toBe(false)
    expect(options.scaleFactor).toBe(100)
  })
})

describe('printToPdfOptionsFromDialog', () => {
  it('uses printToPDF units and a 1-based page range string', () => {
    expect(printToPdfOptionsFromDialog(CHOICE)).toEqual({
      landscape: false,
      pageSize: 'Letter',
      margins: { top: 1, bottom: 1, left: 0.5, right: 0.5 },
      scale: 1,
      printBackground: true,
      pageRanges: '2-4',
    })
  })
})
