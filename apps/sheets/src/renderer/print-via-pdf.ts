/**
 * Print a PDF (base64) with the same path as the PDF app: pdf.js rasterize +
 * system print dialog.
 */
import { printPdf } from '@genoffice/pdf-print'
import { GlobalWorkerOptions, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import workerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'

let workerReady = false

function ensureWorker(): void {
  if (workerReady) return
  GlobalWorkerOptions.workerSrc = workerUrl
  workerReady = true
}

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return bytes
}

export async function printPdfBase64(base64: string, pages?: number[]): Promise<void> {
  ensureWorker()
  const task = getDocument({ data: decodeBase64(base64), useWasm: false })
  const doc = await task.promise
  try {
    await printPdf(doc, pages)
  } finally {
    await task.destroy()
  }
}
