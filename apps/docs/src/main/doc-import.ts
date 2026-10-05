import { mkdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { randomUUID } from 'node:crypto'
import { pathToFileURL } from 'node:url'
import { docTextToHtml, docToText } from '@genoffice/file-parse'
import { convertHtmlToDocx } from '../../../../packages/html2docx/src'
import { ElectronBrowserDriver } from '../../../../packages/html2docx/src/drivers/electron'

/**
 * Legacy .doc import: a Word 97-2003 file has no OOXML structure to patch, so
 * the open path converts it once into a real .docx copy in a temp directory and
 * the session opens that copy (the original stays untouched; the first save
 * routes through Save As, defaulting to a .docx sibling of the original).
 * Text and paragraph structure survive; images and legacy binary formatting
 * have no reader in the suite. The docx assembly reuses the same html2docx
 * chain as the altChunk path: the extracted text is wrapped in plain paragraphs
 * and rendered through the hidden Chromium window.
 */

const IMPORT_VIEWPORT = { width: 794, height: 1123, deviceScaleFactor: 2 }
const IMPORT_TIMEOUT_MS = 120_000

/** Extract a legacy Word 97-2003 document's text and rebuild it as .docx bytes. */
export async function convertDocBytesToDocx(bytes: Uint8Array): Promise<Uint8Array> {
  const text = await docToText(bytes)
  const html = docTextToHtml(text)
  const workDir = join(tmpdir(), `genoffice-doc-import-${randomUUID()}`)
  await mkdir(workDir, { recursive: true })
  let driver: ElectronBrowserDriver | null = null
  try {
    const htmlPath = join(workDir, 'import.html')
    // the BOM outranks a stale <meta charset> left in the decoded markup
    await writeFile(htmlPath, `\ufeff${html}`, 'utf8')
    driver = await ElectronBrowserDriver.create(IMPORT_VIEWPORT)
    // An extraction artifact could in principle keep the renderer busy forever;
    // race a watchdog and destroy the window on timeout, matching the altChunk
    // conversion guard.
    const conversion = convertHtmlToDocx({ url: pathToFileURL(htmlPath).href }, driver, {
      naturalTableWidth: true,
    }).then(({ docx }) => docx)
    let watchdog: ReturnType<typeof setTimeout> | undefined
    const docx = await Promise.race([
      conversion,
      new Promise<never>((_resolve, reject) => {
        watchdog = setTimeout(() => {
          if (driver && !driver.isWindowDestroyed()) driver.destroyNow()
          driver = null
          reject(new Error('legacy .doc conversion timed out'))
        }, IMPORT_TIMEOUT_MS)
      }),
    ]).finally(() => clearTimeout(watchdog))
    return docx
  } finally {
    await driver?.close()
    await rm(workDir, { recursive: true, force: true }).catch(() => {})
  }
}

/**
 * Write the converted .docx into an app-owned temp directory shaped like the
 * sheets .xls import, and report the open path. The caller remembers `tempDir`
 * so the session teardown can remove it.
 */
export async function writeDocImportCopy(
  originalPath: string,
  docx: Uint8Array,
): Promise<{ openPath: string; tempDir: string }> {
  const stem =
    originalPath
      .replace(/\.[^.]+$/, '')
      .split(/[\\/]/)
      .pop() ?? 'document'
  const tempDir = join(tmpdir(), 'genoffice-imports', randomUUID())
  await mkdir(tempDir, { recursive: true })
  const openPath = join(tempDir, `${stem}.docx`)
  await writeFile(openPath, docx)
  return { openPath, tempDir }
}
