import { describe, expect, it } from 'vitest'
import { pptxToText } from '@genoffice/file-parse'
import { buildPptxFromSlideTexts, splitSlideTitleBody } from '../src/main/ppt-import'

describe('splitSlideTitleBody', () => {
  it('takes the first line as the title and the rest as the body', () => {
    expect(splitSlideTitleBody('Q3 Review\nRevenue is up\nCosts are flat')).toEqual({
      title: 'Q3 Review',
      bodyLines: ['Revenue is up', 'Costs are flat'],
    })
  })

  it('keeps a lone short line as a title with no body', () => {
    expect(splitSlideTitleBody('Appendix')).toEqual({ title: 'Appendix', bodyLines: [] })
  })

  it('treats a long first line as body prose, not a title', () => {
    const longLine = 'x'.repeat(200)
    expect(splitSlideTitleBody(`${longLine}\nmore`)).toEqual({
      title: '',
      bodyLines: [longLine, 'more'],
    })
  })
})

describe('buildPptxFromSlideTexts', () => {
  it('rebuilds one slide per source slide, in order, with the extracted text', async () => {
    const bytes = await buildPptxFromSlideTexts([
      'Q3 Review\nRevenue is up\nCosts are flat',
      'Appendix',
    ])
    const text = await pptxToText(bytes)
    // pins the addBlankSlide insert-after semantics: slide order must follow the
    // source text order, and slide 2 must not inherit slide 1's content
    const slide2 = text.slice(text.indexOf('## Slide 2'))
    expect(text).toContain('## Slide 1')
    expect(text).toContain('Q3 Review')
    expect(text).toContain('Revenue is up')
    expect(slide2).toContain('Appendix')
    expect(slide2).not.toContain('Q3 Review')
  })

  it('keeps a textless source slide as an empty slide', async () => {
    const bytes = await buildPptxFromSlideTexts(['Real content', ''])
    const text = await pptxToText(bytes)
    expect(text.match(/## Slide \d/g)).toHaveLength(2)
    expect(text).toContain('Real content')
  })

  it('always produces at least the template slide', async () => {
    const bytes = await buildPptxFromSlideTexts([])
    const text = await pptxToText(bytes)
    expect(text).toContain('## Slide 1')
  })
})
