import { docTextToHtml, docToText, pptToSlideTexts } from '@genoffice/file-parse'
import { buildPptxFromSlideTexts } from '../../../../apps/slides/src/main/ppt-import'
import { readInput } from '../fs'
import { blankDocument, closeDocument, fillFromHtml, saveDocument } from './docx'

/**
 * Legacy 97-2003 import routes for `genoffice convert`: .doc → .docx and
 * .ppt → .pptx. These reuse the same converters as the apps' open paths —
 * text extraction from @genoffice/file-parse, the restricted-HTML → docx
 * pipeline (jsdom-hosted docs editor protocol), and the Electron-free
 * text → pptx builder the slides app imports on open.
 */

/** .doc → .docx: extract the text and rebuild it as native OOXML paragraphs. */
export async function convertLegacyDocToDocx(
  input: string,
): Promise<{ bytes: Uint8Array; detail: Record<string, unknown> }> {
  const text = await docToText(new Uint8Array(readInput(input)))
  const doc = await blankDocument()
  try {
    const blocks = fillFromHtml(doc, docTextToHtml(text))
    return { bytes: await saveDocument(doc), detail: { blocks } }
  } finally {
    closeDocument(doc)
  }
}

/** .ppt → .pptx: one rebuilt slide per source slide, extracted text as title + body boxes. */
export async function convertLegacyPptToPptx(
  input: string,
): Promise<{ bytes: Uint8Array; detail: Record<string, unknown> }> {
  const texts = await pptToSlideTexts(new Uint8Array(readInput(input)))
  return { bytes: await buildPptxFromSlideTexts(texts), detail: { slides: texts.length } }
}
