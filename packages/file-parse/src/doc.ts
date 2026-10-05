import WordExtractor from 'word-extractor'
import './vendor'

function normalizeText(text: string): string {
  return text.replace(/\r\n?/g, '\n').trim()
}

/** Extract readable body and text-box content from a legacy Word 97-2003 file. */
export async function docToText(bytes: Uint8Array): Promise<string> {
  const document = await new WordExtractor().extract(Buffer.from(bytes))
  const body = document.getBody({ filterUnicode: false })
  const textboxes = document.getTextboxes({
    filterUnicode: false,
    includeHeadersAndFooters: false,
  })
  return normalizeText([body, textboxes].filter((part) => part.trim()).join('\n'))
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/**
 * Wrap extracted plain text as one <p> per non-empty line: the restricted-HTML
 * fragment both legacy .doc import paths (docs app, CLI convert) feed into
 * their html→docx machinery. Pure; unit-tested.
 */
export function docTextToHtml(text: string): string {
  const lines = text.split(/\r\n|\r|\n/)
  const paragraphs = lines.map((line) => line.trim()).filter(Boolean)
  const body: string[] = paragraphs.map((p) => `<p>${escapeHtml(p)}</p>`)
  if (body.length === 0) body.push('<p></p>')
  return `<!doctype html><html><body>${body.join('')}</body></html>`
}
