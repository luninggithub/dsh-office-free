import type { Volatile } from '@deepseek-ai/cordis'
import { isAbsolute, join, resolve } from 'node:path'
import z from '@deepseek-ai/schemastery'
import { resolveDshHome } from './dsh-home.ts'
import type { ImageOutputSize } from './service/types.ts'

/** Volcengine Ark image generation endpoint used when none is configured. */
export const DEFAULT_IMAGE_ENDPOINT = 'https://ark.cn-beijing.volces.com/api/v3/images/generations'

/** Image model used when none is configured; the Ark account must have it activated. */
export const DEFAULT_IMAGE_MODEL = 'doubao-seedream-5-0-pro-260628'

/** Output size classes the image model accepts. */
const IMAGE_OUTPUT_SIZES = ['1K', '2K', '4K'] as const

/** Configuration shared by the Univer service provider and its consumers. */
export interface Config {
  /** Initial loopback port used by the bundled Gateway; occupied ports advance by one. */
  gatewayPort?: number
  /** Start the bundled Gateway when file state is first requested. */
  autoStartGateway?: boolean
  /** Maximum time allowed for the bundled Gateway to become healthy. */
  gatewayStartupTimeoutMs?: number
  /** HTTP timeout used for Gateway state reads. */
  gatewayRequestTimeoutMs?: number
  /** HTTP timeout used for Gateway mutations. */
  gatewayMutationTimeoutMs?: number
  /** Maximum time for one content operation, or for opening a pooled runtime worker. */
  unitContentOperationTimeoutMs?: number
  /**
   * Pooled Collaboration Runtime Workers kept for inspection, execution, and render-source
   * reads. `0` runs every operation one-shot.
   */
  unitContentRuntimeMaxEntries?: number
  /** Idle time after which a pooled Collaboration Runtime Worker exits. */
  unitContentRuntimeIdleTtlMs?: number
  /** Maximum lifetime of one browser-backed screenshot operation. */
  screenshotOperationTimeoutMs?: number
  /** Maximum lifetime of one browser-backed PDF print operation. */
  printPdfOperationTimeoutMs?: number
  /** Maximum number of document or slide pages captured by one screenshot call. */
  screenshotMaxPages?: number
  /** Maximum pixel count for each rendered screenshot image. */
  screenshotMaxPixels?: number
  /**
   * Browser executable used by all browser-backed render operations. When unset,
   * the render runtime auto-detects Chrome/Chromium (and Edge on macOS); this
   * option overrides every other detection step, including `UNIVER_RENDER_BROWSER`.
   */
  browserExecutablePath?: string
  /** Persistent cache directory for downloaded resource-library SVGs. */
  resourceCacheRoot?: string
  /** Maximum time allowed for one resource-library download. */
  resourceDownloadTimeoutMs?: number
  /** Maximum lifetime of one resource-library tool operation. */
  resourceOperationTimeoutMs?: number
  /** Maximum time to wait for a collaboration commit acknowledgement before confirming by pull. */
  unitContentCommitTimeoutMs?: number
  /** Freshness window for file state reads. */
  stateCacheTtlMs?: number
  /** Freshness window for changed-unit reads. */
  unitCacheTtlMs?: number
  /** Register model-facing `univer_*` tools. */
  tools?: boolean
  /** Register version-matched bundled Univer skills. */
  skills?: boolean
  /** Send anonymous product telemetry; `DO_NOT_TRACK` also turns it off. */
  telemetry?: boolean
  /**
   * Image model used to edit the image a user selects in the right panel. Only a model the
   * deployment's Ark account has activated can answer, so this must stay configurable.
   */
  imageModel?: string
  /** Output size requested from the image model. */
  imageSize?: ImageOutputSize
  /** Ark API key; `ARK_API_KEY` in the environment is used when this is unset. */
  imageApiKey?: string | Volatile<string | undefined>
  /** Absolute HTTP(S) image generation endpoint; defaults to the Volcengine Ark route. */
  imageEndpoint?: string
  /**
   * Maximum lifetime of one image generation request, including the result download. Ark answers
   * in around a minute for a real edit, so the default leaves room for a slow model turn.
   */
  imageOperationTimeoutMs?: number
}

/** Fully resolved configuration used by the implementation. */
export interface ResolvedConfig {
  readonly gatewayPort: number
  readonly autoStartGateway: boolean
  readonly gatewayStartupTimeoutMs: number
  readonly gatewayRequestTimeoutMs: number
  readonly gatewayMutationTimeoutMs: number
  readonly unitContentOperationTimeoutMs: number
  readonly unitContentRuntimeMaxEntries: number
  readonly unitContentRuntimeIdleTtlMs: number
  readonly screenshotOperationTimeoutMs: number
  readonly printPdfOperationTimeoutMs: number
  readonly screenshotMaxPages: number
  readonly screenshotMaxPixels: number
  readonly browserExecutablePath?: string
  readonly resourceCacheRoot: string
  readonly resourceDownloadTimeoutMs: number
  readonly resourceOperationTimeoutMs: number
  readonly unitContentCommitTimeoutMs: number
  readonly stateCacheTtlMs: number
  readonly unitCacheTtlMs: number
  readonly tools: boolean
  readonly skills: boolean
  readonly telemetry: boolean
  readonly imageModel: string
  readonly imageSize: ImageOutputSize
  /** The key for the next image edit: the live profile value, then `ARK_API_KEY`. */
  readonly readImageApiKey: () => string | undefined
  readonly imageEndpoint: string
  readonly imageOperationTimeoutMs: number
}

/**
 * Cordis configuration schema. The image key is a volatile secret, so its runtime value is read
 * live and this schema is not forced onto the plain input interface.
 */
export const Config = z.object({
  gatewayPort: z.natural().max(65535).default(9080),
  autoStartGateway: z.boolean().default(true),
  gatewayStartupTimeoutMs: z.natural().default(10_000),
  gatewayRequestTimeoutMs: z.natural().default(3_000),
  gatewayMutationTimeoutMs: z.natural().default(60_000),
  unitContentOperationTimeoutMs: z.natural().default(120_000),
  unitContentRuntimeMaxEntries: z.natural().default(20),
  unitContentRuntimeIdleTtlMs: z.natural().default(300_000),
  screenshotOperationTimeoutMs: z.natural().default(120_000),
  printPdfOperationTimeoutMs: z.natural().default(120_000),
  screenshotMaxPages: z.natural().default(30),
  screenshotMaxPixels: z.natural().default(16_777_216),
  browserExecutablePath: z.string(),
  resourceCacheRoot: z.string(),
  resourceDownloadTimeoutMs: z.natural().default(15_000),
  resourceOperationTimeoutMs: z.natural().default(120_000),
  unitContentCommitTimeoutMs: z.natural().default(5_000),
  stateCacheTtlMs: z.natural().default(1_000),
  unitCacheTtlMs: z.natural().default(5_000),
  tools: z.boolean().default(true),
  skills: z.boolean().default(true),
  telemetry: z.boolean().default(true),
  imageModel: z.string(),
  imageSize: z.union(IMAGE_OUTPUT_SIZES),
  imageApiKey: z.string().role('secret').volatile(),
  imageEndpoint: z.string(),
  imageOperationTimeoutMs: z.natural().default(180_000)
})

/** Apply defaults and reject configuration that cannot run. */
export function resolveConfig(config: Config = {}): ResolvedConfig {
  const browserExecutablePath = resolveBrowserExecutablePath(config.browserExecutablePath)
  const resolved: ResolvedConfig = {
    gatewayPort: config.gatewayPort ?? 9080,
    autoStartGateway: config.autoStartGateway ?? true,
    gatewayStartupTimeoutMs: config.gatewayStartupTimeoutMs ?? 10_000,
    gatewayRequestTimeoutMs: config.gatewayRequestTimeoutMs ?? 3_000,
    gatewayMutationTimeoutMs: config.gatewayMutationTimeoutMs ?? 60_000,
    unitContentOperationTimeoutMs: config.unitContentOperationTimeoutMs ?? 120_000,
    unitContentRuntimeMaxEntries: config.unitContentRuntimeMaxEntries ?? 20,
    unitContentRuntimeIdleTtlMs: config.unitContentRuntimeIdleTtlMs ?? 300_000,
    screenshotOperationTimeoutMs: config.screenshotOperationTimeoutMs ?? 120_000,
    printPdfOperationTimeoutMs: config.printPdfOperationTimeoutMs ?? 120_000,
    screenshotMaxPages: config.screenshotMaxPages ?? 30,
    screenshotMaxPixels: config.screenshotMaxPixels ?? 16_777_216,
    ...(browserExecutablePath === undefined ? {} : { browserExecutablePath }),
    resourceCacheRoot: resolveResourceCacheRoot(config.resourceCacheRoot),
    resourceDownloadTimeoutMs: config.resourceDownloadTimeoutMs ?? 15_000,
    resourceOperationTimeoutMs: config.resourceOperationTimeoutMs ?? 120_000,
    unitContentCommitTimeoutMs: config.unitContentCommitTimeoutMs ?? 5_000,
    stateCacheTtlMs: config.stateCacheTtlMs ?? 1_000,
    unitCacheTtlMs: config.unitCacheTtlMs ?? 5_000,
    tools: config.tools ?? true,
    skills: config.skills ?? true,
    telemetry: config.telemetry ?? true,
    imageModel: resolveImageModel(config.imageModel),
    imageSize: resolveImageSize(config.imageSize),
    readImageApiKey: () => resolveImageApiKey(currentImageApiKey(config.imageApiKey)),
    imageEndpoint: resolveImageEndpoint(config.imageEndpoint),
    imageOperationTimeoutMs: config.imageOperationTimeoutMs ?? 180_000
  }
  if (
    !Number.isInteger(resolved.gatewayPort) ||
    resolved.gatewayPort < 1 ||
    resolved.gatewayPort > 65_535
  ) {
    throw new Error('univer: gatewayPort must be an integer between 1 and 65535')
  }
  for (const [name, value] of Object.entries({
    gatewayStartupTimeoutMs: resolved.gatewayStartupTimeoutMs,
    gatewayRequestTimeoutMs: resolved.gatewayRequestTimeoutMs,
    gatewayMutationTimeoutMs: resolved.gatewayMutationTimeoutMs,
    unitContentOperationTimeoutMs: resolved.unitContentOperationTimeoutMs,
    unitContentRuntimeIdleTtlMs: resolved.unitContentRuntimeIdleTtlMs,
    screenshotOperationTimeoutMs: resolved.screenshotOperationTimeoutMs,
    printPdfOperationTimeoutMs: resolved.printPdfOperationTimeoutMs,
    screenshotMaxPages: resolved.screenshotMaxPages,
    screenshotMaxPixels: resolved.screenshotMaxPixels,
    resourceDownloadTimeoutMs: resolved.resourceDownloadTimeoutMs,
    resourceOperationTimeoutMs: resolved.resourceOperationTimeoutMs,
    unitContentCommitTimeoutMs: resolved.unitContentCommitTimeoutMs,
    stateCacheTtlMs: resolved.stateCacheTtlMs,
    unitCacheTtlMs: resolved.unitCacheTtlMs,
    imageOperationTimeoutMs: resolved.imageOperationTimeoutMs
  })) {
    if (!Number.isSafeInteger(value) || value < 1)
      throw new Error(`univer: ${name} must be a positive integer`)
  }
  if (
    !Number.isSafeInteger(resolved.unitContentRuntimeMaxEntries) ||
    resolved.unitContentRuntimeMaxEntries < 0
  ) {
    throw new Error('univer: unitContentRuntimeMaxEntries must be a non-negative integer')
  }
  return resolved
}

function resolveResourceCacheRoot(configured: string | undefined): string {
  if (configured !== undefined) {
    if (configured.trim().length === 0 || !isAbsolute(configured)) {
      throw new Error('univer: resourceCacheRoot must be a non-empty absolute path')
    }
    return resolve(configured)
  }
  return join(resolveDshHome(), 'cache', 'dsh-univer-office', 'resources')
}

/** A relative browser path would silently depend on the Host process working directory. */
function resolveBrowserExecutablePath(configured: string | undefined): string | undefined {
  if (configured === undefined) return undefined
  const trimmed = configured.trim()
  if (trimmed.length === 0 || !isAbsolute(trimmed)) {
    throw new Error('univer: browserExecutablePath must be a non-empty absolute path')
  }
  return resolve(trimmed)
}

function resolveImageModel(configured: string | undefined): string {
  const model = configured?.trim() ?? DEFAULT_IMAGE_MODEL
  if (model.length === 0) throw new Error('univer: imageModel must not be empty')
  return model
}

function resolveImageSize(configured: string | undefined): ImageOutputSize {
  const size = (configured ?? '1K').trim().toUpperCase()
  if (size === '1K' || size === '2K' || size === '4K') return size
  throw new Error('univer: imageSize must be one of 1K, 2K, or 4K')
}

/** A schema-wrapped key stays live; tests and direct callers may still pass the string. */
function currentImageApiKey(
  configured: string | Volatile<string | undefined> | undefined
): string | undefined {
  if (typeof configured === 'object' && configured !== null && 'get' in configured) {
    return configured.get()
  }
  return configured
}

/** Blank means "unset", so an empty environment variable cannot masquerade as a key. */
function resolveImageApiKey(configured: string | undefined): string | undefined {
  const key = (configured ?? process.env.ARK_API_KEY ?? '').trim()
  return key.length === 0 ? undefined : key
}

function resolveImageEndpoint(configured: string | undefined): string {
  const endpoint = (configured ?? DEFAULT_IMAGE_ENDPOINT).trim()
  let url: URL
  try {
    url = new URL(endpoint)
  } catch {
    throw new Error('univer: imageEndpoint must be an absolute HTTP(S) URL')
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') {
    throw new Error('univer: imageEndpoint must use http or https')
  }
  return endpoint
}
