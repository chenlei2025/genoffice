import { execFile } from 'node:child_process'

/**
 * macOS `webContents.print()` fails closed when Chromium cannot enumerate a
 * printer ("no valid printers available" / "No printers available on the
 * network") and when the target window is hidden. The system print panel,
 * opened on a PDF, still offers every configured printer and Save as PDF.
 *
 * The path is an argv element, never interpolated into the script.
 */
const MAC_PRINT_PANEL_SCRIPT = `function run(argv) {
  ObjC.import('Quartz')
  ObjC.import('AppKit')
  var path = argv[0]
  if (!path) return 'error:missing path'
  var doc = $.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath($(path)))
  if (doc.isNil()) return 'error:unreadable'
  var pages = Number(ObjC.unwrap(doc.pageCount))
  if (!(pages > 0)) return 'error:unreadable'
  // 2 = scale down to fit. The PDF is already the document's page size, so it
  // is not enlarged; a smaller printer sheet still gets the whole page.
  var op = doc.printOperationForPrintInfoScalingModeAutoRotate($.NSPrintInfo.sharedPrintInfo, 2, true)
  op.setShowsPrintPanel(true)
  op.setShowsProgressPanel(true)
  return op.runOperation() ? 'ok' : 'canceled'
}`

export type MacPrintPanelResult = { ok: true } | { ok: false; error?: string }

/** `ok: false` with no error is the user closing the panel. */
export function interpretMacPrintPanelResult(stdout: string): MacPrintPanelResult {
  const value = stdout.trim()
  if (value === 'ok') return { ok: true }
  if (value === 'canceled') return { ok: false }
  const detail = value.replace(/^error:/, '').trim()
  return { ok: false, error: detail || 'print failed' }
}

export type RunCommand = (command: string, args: string[]) => Promise<string>

function execFileText(command: string, args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(command, args, { timeout: 0, windowsHide: true }, (err, stdout, stderr) => {
      if (err) reject(new Error(String(stderr || err.message).trim() || 'print failed'))
      else resolve(String(stdout))
    })
  })
}

export function presentMacPrintPanel(
  pdfPath: string,
  run: RunCommand = execFileText,
): Promise<MacPrintPanelResult> {
  return run('osascript', ['-l', 'JavaScript', '-e', MAC_PRINT_PANEL_SCRIPT, '--', pdfPath]).then(
    (stdout) => interpretMacPrintPanelResult(stdout),
    (err: unknown) => ({
      ok: false as const,
      error: err instanceof Error ? err.message : String(err),
    }),
  )
}
