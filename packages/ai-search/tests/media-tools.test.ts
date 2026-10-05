import { describe, expect, it, vi, beforeEach } from 'vitest'

const generateWithProvider = vi.fn()
const analyzeWithProvider = vi.fn()
vi.mock('@genoffice/ai-provider', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@genoffice/ai-provider')>()),
  generateImageWithProvider: (...args: unknown[]) => generateWithProvider(...args),
  analyzeMediaWithProvider: (...args: unknown[]) => analyzeWithProvider(...args),
}))

import { analyzeMediaTool, generateImageTool, type GenerateImageToolOp } from '../src/media-tools'

// nonexistent settings file → defaults: no media provider configured
const SETTINGS = '/nonexistent/ai-settings.json'

beforeEach(() => {
  generateWithProvider.mockReset()
  analyzeWithProvider.mockReset()
})

describe('generateImageTool routing', () => {
  it('reports an unconfigured capability instead of attempting a request', async () => {
    const r = await generateImageTool(SETTINGS, { prompt: 'red podcast icon' })
    expect(r.error).toMatch(/No image provider is configured/)
    expect(generateWithProvider).not.toHaveBeenCalled()
  })

  it('surfaces provider failures as tool errors', async () => {
    // point the settings file at a configured openai image provider
    const { mkdtempSync, writeFileSync } = await import('node:fs')
    const { tmpdir } = await import('node:os')
    const { join } = await import('node:path')
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-media-tools-'))
    const path = join(dir, 'ai-settings.json')
    writeFileSync(
      path,
      // the top-level providers block must exist: a settings file without one
      // is treated as pre-provider legacy and media is reset to defaults
      JSON.stringify({
        provider: 'openai',
        providers: {},
        media: {
          imageProvider: 'openai',
          providers: { openai: { apiKey: 'k', imageModel: 'gpt-image-2' } },
        },
      }),
    )
    generateWithProvider.mockRejectedValueOnce(new Error('quota exceeded'))
    const op: GenerateImageToolOp = { prompt: 'red podcast icon' }
    const r = await generateImageTool(path, op)
    expect(r).toEqual({ error: 'quota exceeded' })
    expect(generateWithProvider).toHaveBeenCalledTimes(1)
  })
})

describe('analyzeMediaTool routing', () => {
  it('reports an unconfigured capability instead of attempting a request', async () => {
    const r = await analyzeMediaTool(SETTINGS, {
      mediaUrls: ['https://cdn/x/pic.png'],
      requirements: 'describe these',
    })
    expect(r.error).toMatch(/No media analysis provider is configured/)
    expect(analyzeWithProvider).not.toHaveBeenCalled()
  })
})
