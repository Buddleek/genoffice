import { join } from 'node:path'
import { genofficeUserDataDir } from './gui'

/**
 * The cloud commands (search / image / media) reuse the editors' provider
 * routing: the BYOK providers configured in the app's AI settings. That
 * settings file lives in the shell's Electron userData directory, which
 * genoffice has to locate without Electron.
 */
export function aiSettingsPath(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_AI_SETTINGS || join(genofficeUserDataDir(env), 'ai-settings.json')
}
