import {
  activeMediaProvider,
  activeSearchProvider,
  imageGenerationAvailable,
  mediaAnalysisAvailable,
} from '@genoffice/ai-provider'
import { readAiSettingsFile } from '@genoffice/ai-search'
import { aiSettingsPath } from '../cloud'
import type { CommandDef } from '../registry'
import { appLaunch } from '../resources'

/**
 * What the cloud commands can do on this machine, decided from GenOffice's
 * own settings without a network call: a keyed search provider (or Parallel,
 * which also answers keylessly), or a BYOK media key. Unkeyed fallbacks in the
 * default chain (free Parallel MCP, DuckDuckGo) do not count as configured for
 * image search, but keyless Parallel web search does. Agents check this once
 * before planning work that needs photos or web facts.
 */
export const capabilitiesCommand: CommandDef = {
  name: 'capabilities',
  summary:
    'Report which cloud features (search, image search, image generation, media analysis) are configured in GenOffice, and whether the app is installed.',
  usage: 'capabilities',
  async run(_args, ctx) {
    const settings = readAiSettingsFile(aiSettingsPath(ctx.env))
    const searchProvider = activeSearchProvider(settings)
    const imageSearch = searchProvider === 'serper' || searchProvider === 'serply'
    const imageGeneration = imageGenerationAvailable(settings)
    const mediaAnalysis = mediaAnalysisAvailable(settings)
    const detail = {
      search: {
        // keyless Parallel counts: the free Search MCP needs no key
        available: searchProvider !== null,
        via: searchProvider,
      },
      image_search: {
        available: imageSearch,
        via: imageSearch ? searchProvider : null,
      },
      image_generation: {
        available: imageGeneration,
        via: imageGeneration ? activeMediaProvider(settings, 'image') : null,
      },
      media_analysis: {
        available: mediaAnalysis,
        via: mediaAnalysis ? activeMediaProvider(settings, 'analysis') : null,
      },
      app: { available: appLaunch(ctx.env) !== null },
      settings_path: aiSettingsPath(ctx.env),
    }
    const on = Object.entries(detail)
      .filter(([k, v]) => k !== 'settings_path' && (v as { available: boolean }).available)
      .map(([k]) => k)
    return {
      summary: on.length
        ? `configured: ${on.join(', ')}`
        : 'no cloud feature configured; the app is not installed',
      detail,
    }
  },
}
