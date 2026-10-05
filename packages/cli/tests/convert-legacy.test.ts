import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { docxToText, pptxToText } from '@genoffice/file-parse'
import { convertLegacyDocToDocx, convertLegacyPptToPptx } from '../src/formats/legacy'

function legacyFixture(name: string): string {
  return fileURLToPath(new URL(`../../file-parse/tests/fixtures/${name}`, import.meta.url))
}

async function withTempLegacyFile<T>(
  name: string,
  bytes: Buffer,
  fn: (path: string) => Promise<T>,
): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'convert-legacy-'))
  try {
    const path = join(dir, name)
    await writeFile(path, bytes)
    return await fn(path)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

describe('convertLegacyDocToDocx', () => {
  it('rebuilds a Word 97-2003 file as a real .docx with the extracted paragraphs', async () => {
    const { bytes, detail } = await convertLegacyDocToDocx(legacyFixture('legacy-sample.doc'))
    expect(bytes.byteLength).toBeGreaterThan(0)
    expect(Number(detail.blocks)).toBeGreaterThan(0)
    const text = await docxToText(bytes)
    expect(text).toContain('Legacy Report')
    expect(text).toContain('Legacy DOC body text')
    expect(text).toContain('Second paragraph from Word 97-2003.')
  })

  it('fails with an error on a corrupt .doc', async () => {
    await expect(
      withTempLegacyFile('corrupt.doc', Buffer.from('not a legacy document'), (path) =>
        convertLegacyDocToDocx(path),
      ),
    ).rejects.toThrow()
  })
})

describe('convertLegacyPptToPptx', () => {
  it('rebuilds a PowerPoint 97-2003 file as a real .pptx, one slide per source slide', async () => {
    const { bytes, detail } = await convertLegacyPptToPptx(legacyFixture('legacy-sample.ppt'))
    expect(bytes.byteLength).toBeGreaterThan(0)
    expect(detail.slides).toBe(2)
    const text = await pptxToText(bytes)
    expect(text).toContain('Legacy PPT title')
    expect(text).toContain('First slide body')
    const slide2 = text.slice(text.indexOf('## Slide 2'))
    expect(slide2).toContain('Second legacy slide')
    expect(slide2).not.toContain('Legacy PPT title')
  })

  it('fails with an error on a corrupt .ppt', async () => {
    await expect(
      withTempLegacyFile('corrupt.ppt', Buffer.from('not a legacy document'), (path) =>
        convertLegacyPptToPptx(path),
      ),
    ).rejects.toThrow()
  })
})
