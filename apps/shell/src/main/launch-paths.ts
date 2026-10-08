import { existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const SUPPORTED_LAUNCH_RE = /\.(docx|xlsx|xlsm|xls|csv|tsv|pptx|pdf|md|markdown|html?)$/i
const UNSUPPORTED_LAUNCH_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsb|pages|key|numbers)$/i

/**
 * Finder and file associations pass a filesystem path. WeChat's system-open
 * (Qt `QDesktopServices` / `NSWorkspace`) sometimes passes a `file://` URL,
 * with percent-encoding and surrounding quotes. `existsSync` rejects those,
 * so the document is dropped and the app looks like it never launched.
 */
export function normalizeLaunchPath(raw: string): string {
  let value = raw.trim().replace(/[\u0000\r\n]+$/g, '')
  if (value.length >= 2) {
    const quote = value[0]
    if ((quote === '"' || quote === "'") && value.endsWith(quote)) value = value.slice(1, -1)
  }
  if (/^file:/i.test(value)) {
    try {
      value = fileURLToPath(value)
    } catch {
      // not a usable file URL; the exists() check below drops it
    }
  }
  return value
}

interface LaunchPayload {
  launchPaths?: unknown
  launchPath?: unknown
}

function payloadLaunchPaths(additionalData: unknown): string[] {
  if (!additionalData || typeof additionalData !== 'object') return []
  const data = additionalData as LaunchPayload
  const paths = Array.isArray(data.launchPaths)
    ? data.launchPaths.filter((path): path is string => typeof path === 'string' && path.length > 0)
    : []
  if (typeof data.launchPath === 'string' && data.launchPath.length > 0) paths.push(data.launchPath)
  return paths.map(normalizeLaunchPath).filter((path) => path.length > 0)
}

export function collectLaunchPaths(
  argv: readonly string[],
  additionalData?: unknown,
  exists: (path: string) => boolean = existsSync,
): string[] {
  const args = argv.map(normalizeLaunchPath)
  const supported = args.filter((path) => SUPPORTED_LAUNCH_RE.test(path) && exists(path))
  const payload = payloadLaunchPaths(additionalData)
  const fallback =
    supported.length === 0 && payload.length === 0
      ? args.find((path) => UNSUPPORTED_LAUNCH_RE.test(path) && exists(path))
      : undefined
  return [...new Set([...supported, ...payload, ...(fallback ? [fallback] : [])])]
}
