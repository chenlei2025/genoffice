import { FULL_LOAD_MAX_CELLS } from './app-constants'

export type PrintLoadPlan = 'ready' | 'wait' | 'load' | 'too-large'

export interface PrintLoadExtent {
  readonly rows: number
  readonly columns: number
}

export interface PrintLoadRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/**
 * Streaming workbooks never finish `preloadComplete` on their own, so a print
 * that only waits never starts. Load the active sheet when it fits; a sheet
 * already covering its file extent (or a sheet that isn't streamed) can print
 * immediately. `wait` means a preload is already running.
 */
export function printSheetLoadPlan(input: {
  readonly preloadComplete: boolean
  readonly preloadRunning: boolean
  readonly extent: PrintLoadExtent | null
  readonly loaded: PrintLoadRange | null
}): PrintLoadPlan {
  if (input.preloadComplete) return 'ready'
  const extent = input.extent
  if (!extent || extent.rows < 1 || extent.columns < 1) return 'ready'
  const loaded = input.loaded
  if (
    loaded &&
    loaded.startRow <= 0 &&
    loaded.startColumn <= 0 &&
    loaded.endRow >= extent.rows - 1 &&
    loaded.endColumn >= extent.columns - 1
  ) {
    return 'ready'
  }
  if (input.preloadRunning) return 'wait'
  if (extent.rows * extent.columns > FULL_LOAD_MAX_CELLS) return 'too-large'
  return 'load'
}
