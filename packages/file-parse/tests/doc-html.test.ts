import { describe, expect, it } from 'vitest'
import { docTextToHtml } from '../src/doc'

describe('docTextToHtml', () => {
  it('wraps each non-empty line as a paragraph', () => {
    expect(docTextToHtml('First line\nSecond line')).toBe(
      '<!doctype html><html><body><p>First line</p><p>Second line</p></body></html>',
    )
  })

  it('skips blank lines instead of emitting empty paragraphs', () => {
    const html = docTextToHtml('One\n\n\nTwo')
    expect(html).toContain('<p>One</p><p>Two</p>')
    expect(html).not.toContain('<p></p>')
  })

  it('normalizes CRLF line endings', () => {
    expect(docTextToHtml('A\r\nB\rC')).toBe(docTextToHtml('A\nB\nC'))
  })

  it('escapes HTML-significant characters', () => {
    const html = docTextToHtml('1 < 2 & "quotes" > \'apostrophe\'')
    expect(html).toContain('<p>1 &lt; 2 &amp; &quot;quotes&quot; &gt; &#39;apostrophe&#39;</p>')
    expect(html).not.toContain('<p>1 < 2')
  })

  it('emits one empty paragraph for textless input so the docx is not empty', () => {
    expect(docTextToHtml('')).toBe('<!doctype html><html><body><p></p></body></html>')
    expect(docTextToHtml('\n \n')).toBe('<!doctype html><html><body><p></p></body></html>')
  })
})
