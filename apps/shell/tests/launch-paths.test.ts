import { describe, expect, it } from 'vitest'

import { collectLaunchPaths, normalizeLaunchPath } from '../src/main/launch-paths'

describe('collectLaunchPaths', () => {
  it('collects supported argv files in order and removes duplicates', () => {
    expect(
      collectLaunchPaths(
        ['GenOffice.exe', 'first.docx', 'notes.txt', 'second.xlsx', 'first.docx'],
        undefined,
        () => true,
      ),
    ).toEqual(['first.docx', 'second.xlsx'])
  })

  it('accepts a .tsv from argv or the second-instance payload', () => {
    expect(
      collectLaunchPaths(['GenOffice.app', '/data/variants.tsv'], undefined, () => true),
    ).toEqual(['/data/variants.tsv'])
    // a bare .txt is still not a document
    expect(collectLaunchPaths(['GenOffice.app', '/notes.txt'], undefined, () => true)).toEqual([])
  })

  it('collects argv and second-instance payload files without duplicates', () => {
    expect(
      collectLaunchPaths(
        ['GenOffice.exe', 'first.docx', 'second.pptx'],
        {
          launchPaths: ['first.docx', 'third.pdf', 42, ''],
          launchPath: 'legacy.md',
        },
        () => true,
      ),
    ).toEqual(['first.docx', 'second.pptx', 'third.pdf', 'legacy.md'])
  })

  it('accepts a legacy launchPath payload', () => {
    expect(
      collectLaunchPaths(['GenOffice.exe'], { launchPath: 'legacy.docx' }, () => true),
    ).toEqual(['legacy.docx'])
  })

  it('accepts a file URL, including percent-encoded and quoted paths', () => {
    expect(normalizeLaunchPath('file:///private/tmp/%E5%BE%AE%E4%BF%A1.docx')).toBe(
      '/private/tmp/微信.docx',
    )
    expect(normalizeLaunchPath('"file://localhost/tmp/report.xlsx"')).toBe('/tmp/report.xlsx')
    expect(
      collectLaunchPaths(
        ['GenOffice', 'file:///data/%E6%8A%A5%E5%91%8A.docx', '"/tmp/notes.xlsx"'],
        undefined,
        () => true,
      ),
    ).toEqual(['/data/报告.docx', '/tmp/notes.xlsx'])
  })

  it('falls back to the first existing unsupported argv file', () => {
    expect(
      collectLaunchPaths(
        ['GenOffice.exe', 'missing.doc', 'legacy.rtf'],
        undefined,
        (path) => path === 'legacy.rtf',
      ),
    ).toEqual(['legacy.rtf'])
  })
})
