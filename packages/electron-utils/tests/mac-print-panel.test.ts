import { describe, expect, it } from 'vitest'

import {
  interpretMacPrintPanelResult,
  presentMacPrintPanel,
} from '../src/mac-print-panel'

describe('interpretMacPrintPanelResult', () => {
  it('treats a closed panel as cancel, not a failure', () => {
    expect(interpretMacPrintPanelResult('ok\n')).toEqual({ ok: true })
    expect(interpretMacPrintPanelResult('canceled\n')).toEqual({ ok: false })
  })

  it('keeps a script error message', () => {
    expect(interpretMacPrintPanelResult('error:unreadable\n')).toEqual({
      ok: false,
      error: 'unreadable',
    })
  })
})

describe('presentMacPrintPanel', () => {
  it('passes the PDF path as an argument, not inside the script', async () => {
    const path = '/tmp/报表 2026.pdf'
    let command = ''
    let args: string[] = []
    const result = await presentMacPrintPanel(path, (cmd, argv) => {
      command = cmd
      args = argv
      return Promise.resolve('canceled\n')
    })
    expect(result).toEqual({ ok: false })
    expect(command).toBe('osascript')
    expect(args.slice(0, 3)).toEqual(['-l', 'JavaScript', '-e'])
    expect(args[3]).not.toContain('报表')
    expect(args.at(-2)).toBe('--')
    expect(args.at(-1)).toBe(path)
  })
})
