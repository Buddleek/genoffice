import { mkdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { randomUUID } from 'node:crypto'
import { createBlankPptx, openPptx, savePptx, type OpenedPptx } from '@genoffice/pptx-engine'
import { runTxn, type Op } from '@genoffice/pptx-ops'

/**
 * Legacy .ppt import: rebuild a 97-2003 presentation as a real .pptx, one slide
 * per source slide, with the extracted text laid out as a title box plus a body
 * box. Formatting beyond text (images, masters, animations) has no legacy
 * reader in the suite, so the import is a faithful text skeleton the user can
 * design on top of — the same trade-off the CLI's markdown→pptx path makes.
 * Electron-free so the CLI convert route can reuse it (see genoffice convert).
 */

/** page margin: 0.5" (EMU) */
const MARGIN_EMU = 457200
/** gap between the title and body boxes: 0.25" */
const TITLE_BODY_GAP_EMU = 228600
const TITLE_FONT_SIZE = 32
const BODY_FONT_SIZE = 18
/** a first line longer than this reads as body prose, not a slide title */
const MAX_TITLE_CHARS = 120

interface SlideBoxes {
  title: string
  bodyLines: string[]
}

/** first line becomes the title (when short enough), remaining lines the body */
export function splitSlideTitleBody(text: string): SlideBoxes {
  const lines = text.split('\n').map((line) => line.trimEnd())
  const first = (lines[0] ?? '').trim()
  if (first && first.length <= MAX_TITLE_CHARS && lines.length > 1) {
    return {
      title: first,
      bodyLines: lines
        .slice(1)
        .map((l) => l.trim())
        .filter(Boolean),
    }
  }
  if (first && first.length <= MAX_TITLE_CHARS && lines.length === 1) {
    return { title: first, bodyLines: [] }
  }
  return { title: '', bodyLines: lines.map((l) => l.trim()).filter(Boolean) }
}

interface BoxSpec {
  kind: 'textbox'
  offset: { x: number; y: number; cx: number; cy: number }
  paragraphs: Array<{ runs: Array<{ text: string; bold?: boolean; fontSize?: number }> }>
}

function slideTextBoxes(text: string, deck: { cx: number; cy: number }): BoxSpec[] {
  const { title, bodyLines } = splitSlideTitleBody(text)
  if (!title && bodyLines.length === 0) return []
  const width = deck.cx - 2 * MARGIN_EMU
  const boxes: BoxSpec[] = []
  if (title) {
    boxes.push({
      kind: 'textbox',
      offset: { x: MARGIN_EMU, y: MARGIN_EMU, cx: width, cy: 914400 },
      paragraphs: [{ runs: [{ text: title, bold: true, fontSize: TITLE_FONT_SIZE }] }],
    })
  }
  if (bodyLines.length > 0) {
    const y = MARGIN_EMU + (title ? 914400 + TITLE_BODY_GAP_EMU : 0)
    boxes.push({
      kind: 'textbox',
      offset: { x: MARGIN_EMU, y, cx: width, cy: Math.max(deck.cy - y - MARGIN_EMU, 914400) },
      paragraphs: bodyLines.map((line) => ({
        runs: [{ text: line, fontSize: BODY_FONT_SIZE }],
      })),
    })
  }
  return boxes
}

/** Build a .pptx from per-slide plain text (see @genoffice/file-parse pptToSlideTexts). */
export async function buildPptxFromSlideTexts(texts: string[]): Promise<Uint8Array> {
  const opened: OpenedPptx = await openPptx(await createBlankPptx())
  const deck = { cx: opened.deck.size.cx, cy: opened.deck.size.cy }
  // Two transactions: ops are validated against the deck state before any op
  // applies, so the addElement ops may only reference slides once the slide
  // additions have landed.
  const slideOps: Op[] = []
  // the blank template ships one slide; grow the deck to the slide count first
  for (let i = 1; i < texts.length; i++) {
    slideOps.push({ op: 'addBlankSlide', target: { slide: i - 1 } } as Op)
  }
  if (slideOps.length > 0) applyOpsAtomic(opened, slideOps)
  const elementOps: Op[] = []
  texts.forEach((text, slide) => {
    for (const box of slideTextBoxes(text, deck)) {
      elementOps.push({ op: 'addElement', target: { slide }, ...box } as Op)
    }
  })
  if (elementOps.length > 0) applyOpsAtomic(opened, elementOps)
  return savePptx(opened)
}

function applyOpsAtomic(opened: OpenedPptx, ops: Op[]): void {
  const result = runTxn(opened, { ops, isolation: 'atomic' })
  if (!result.applied) {
    const failure = result.failures?.[0]
    throw new Error(
      `legacy PPT rebuild failed: ${failure?.error ?? `op ${failure?.index} did not apply`}`,
    )
  }
}

/**
 * Write the converted .pptx into an app-owned temp directory shaped like the
 * sheets .xls import, and report the open path. The caller remembers `tempDir`
 * so the session teardown can remove it.
 */
export async function writePptImportCopy(
  originalPath: string,
  pptx: Uint8Array,
): Promise<{ openPath: string; tempDir: string }> {
  const stem =
    originalPath
      .replace(/\.[^.]+$/, '')
      .split(/[\\/]/)
      .pop() ?? 'presentation'
  const tempDir = join(tmpdir(), 'genoffice-imports', randomUUID())
  await mkdir(tempDir, { recursive: true })
  const openPath = join(tempDir, `${stem}.pptx`)
  await writeFile(openPath, pptx)
  return { openPath, tempDir }
}
