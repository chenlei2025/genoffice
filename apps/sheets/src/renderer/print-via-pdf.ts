/**
 * Sheets → shared PDF-app print path (native PDF viewer + system print dialog).
 */
import { printPdf } from '@genoffice/pdf-print'

function decodeBase64(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return new Uint8Array(bytes)
}

/** Print PDF bytes via Chromium's native viewer (full system print UI). */
export async function printPdfBase64(base64: string, pages?: number[]): Promise<void> {
  await printPdf(decodeBase64(base64), pages)
}
