import { describe, expect, it } from 'vitest'

import { FULL_LOAD_MAX_CELLS } from '../src/renderer/app-constants'
import { printSheetLoadPlan } from '../src/renderer/print-load-plan'

describe('printSheetLoadPlan', () => {
  it('prints a streamed sheet once its grid covers the file extent', () => {
    expect(
      printSheetLoadPlan({
        preloadComplete: false,
        preloadRunning: false,
        extent: { rows: 40, columns: 8 },
        loaded: { startRow: 0, endRow: 39, startColumn: 0, endColumn: 7 },
      }),
    ).toBe('ready')
  })

  it('loads the active sheet instead of waiting forever', () => {
    expect(
      printSheetLoadPlan({
        preloadComplete: false,
        preloadRunning: false,
        extent: { rows: 4000, columns: 12 },
        loaded: { startRow: 0, endRow: 30, startColumn: 0, endColumn: 11 },
      }),
    ).toBe('load')
  })

  it('waits out a preload that is already running', () => {
    expect(
      printSheetLoadPlan({
        preloadComplete: false,
        preloadRunning: true,
        extent: { rows: 100, columns: 4 },
        loaded: null,
      }),
    ).toBe('wait')
  })

  it('refuses a sheet that cannot be installed', () => {
    expect(
      printSheetLoadPlan({
        preloadComplete: false,
        preloadRunning: false,
        extent: { rows: FULL_LOAD_MAX_CELLS + 1, columns: 1 },
        loaded: null,
      }),
    ).toBe('too-large')
  })

  it('does not block a sheet that is not streamed from the file', () => {
    expect(
      printSheetLoadPlan({
        preloadComplete: false,
        preloadRunning: false,
        extent: null,
        loaded: null,
      }),
    ).toBe('ready')
  })
})
