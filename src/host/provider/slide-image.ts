import { UniverError } from '../service/errors.ts'
import type { JsonValue } from '../service/types.ts'

/** What the browser reported: one image element on one slide page. */
export interface SlideImageTarget {
  readonly pageId: string
  readonly elementId: string
}

/** The resolved image element, with the intrinsic size it occupies on the page. */
export interface SlideImageElement extends SlideImageTarget {
  readonly source: string
  readonly sourceType: 'BASE64' | 'UUID' | 'URL' | undefined
  readonly width: number
  readonly height: number
}

/** An image the model can be given as a reference. */
export interface ReferenceImage {
  readonly bytes: Uint8Array
  readonly mediaType: string
}

/** Where the image's bytes can be read from, as the Gateway and the model see them. */
export interface SlideImageReadDeps {
  readonly gateway: string
  readonly fileKey: string
  readonly worktreeId: string
  readonly timeoutMs: number
  readonly signal?: AbortSignal
}

/**
 * Find the reported image element in one Slide Unit's exported data.
 *
 * A selection is browser state that may be stale by the time the model acts, so every step is
 * re-checked here: the page must still exist, the element must still be an image, and it must carry
 * a source this feature can read.
 */
export function locateSlideImage(unitData: JsonValue, target: SlideImageTarget): SlideImageElement {
  const slides = isRecord(unitData) ? unitData.slides : undefined
  const page = isRecord(slides) ? slides[target.pageId] : undefined
  if (!isRecord(page)) {
    throw new UniverError(
      'The selected slide page no longer exists in this presentation.',
      'IMAGE_SELECTION_STALE'
    )
  }
  const elements = isRecord(page.elements) ? page.elements : undefined
  const element = isRecord(elements) ? elements[target.elementId] : undefined
  if (!isRecord(element)) {
    throw new UniverError(
      'The selected element no longer exists on this slide.',
      'IMAGE_SELECTION_STALE'
    )
  }
  if (element.type !== 'image') {
    throw new UniverError(
      'The selected element is not an image. Select an image on a slide and try again.',
      'IMAGE_SELECTION_NOT_IMAGE'
    )
  }
  const source = element.source
  if (typeof source !== 'string' || source.length === 0) {
    throw new UniverError('The selected image has no source.', 'IMAGE_SOURCE_UNSUPPORTED')
  }
  const transform = isRecord(element.transform) ? element.transform : undefined
  return {
    pageId: target.pageId,
    elementId: target.elementId,
    source,
    sourceType: readSourceType(element.imageSourceType),
    width: readDimension(transform?.width),
    height: readDimension(transform?.height)
  }
}

/** Read the reference bytes from a data URI, the Gateway's asset store, or a plain URL. */
export async function readSlideImageBytes(
  element: SlideImageElement,
  deps: SlideImageReadDeps
): Promise<ReferenceImage> {
  if (element.source.startsWith('data:')) return decodeDataUri(element.source)
  if (/^https?:\/\//iu.test(element.source) || element.sourceType === 'URL') {
    const bytes = await fetchBytes(
      deps,
      element.source,
      'The selected image URL could not be read.'
    )
    return { bytes, mediaType: sniffMediaType(bytes) }
  }
  // A stored asset id is the remaining form: the Gateway serves its bytes inside this branch.
  const bytes = await fetchBytes(
    deps,
    assetContentUrl(deps, element.source),
    'The selected image could not be read from the Gateway.'
  )
  return { bytes, mediaType: sniffMediaType(bytes) }
}

/**
 * Facade code that swaps one image's content while keeping its frame.
 *
 * `setSource` replaces the picture in place, so position, size, rotation, and z-order all survive;
 * the caller has already confirmed the element is an image in this Unit.
 */
export function replaceSlideImageCode(target: SlideImageTarget, dataUrl: string): string {
  return [
    `const slide = presentation.getSlideById(${JSON.stringify(target.pageId)});`,
    `if (!slide) throw new Error('Slide page not found');`,
    `const image = slide.getImages().find((candidate) => candidate.getId() === ${JSON.stringify(target.elementId)});`,
    `if (!image) throw new Error('Image element not found');`,
    `image.setSource(${JSON.stringify(dataUrl)}, 'BASE64');`,
    `return { pageId: slide.getId(), elementId: image.getId() };`
  ].join('\n')
}

/** Build the data URI the Facade stores and the Gateway externalizes on commit. */
export function imageDataUrl(image: ReferenceImage): string {
  return `data:${image.mediaType};base64,${Buffer.from(image.bytes).toString('base64')}`
}

function assetContentUrl(deps: SlideImageReadDeps, assetId: string): string {
  const origin = deps.gateway.replace(/\/+$/u, '')
  return (
    `${origin}/uf/${encodeURIComponent(deps.fileKey)}` +
    `/worktrees/${encodeURIComponent(deps.worktreeId)}/universer-api/file/` +
    `${encodeURIComponent(assetId)}/content`
  )
}

async function fetchBytes(
  deps: SlideImageReadDeps,
  url: string,
  fallbackMessage: string
): Promise<Uint8Array> {
  const timeout = AbortSignal.timeout(deps.timeoutMs)
  const signal = deps.signal === undefined ? timeout : AbortSignal.any([deps.signal, timeout])
  let response: Response
  try {
    response = await fetch(url, { signal })
  } catch (error) {
    if (deps.signal?.aborted === true) throw error
    throw new UniverError(fallbackMessage, 'IMAGE_SOURCE_UNAVAILABLE', { cause: error })
  }
  if (!response.ok) {
    throw new UniverError(
      `${fallbackMessage} (HTTP ${String(response.status)})`,
      'IMAGE_SOURCE_UNAVAILABLE'
    )
  }
  const bytes = new Uint8Array(await response.arrayBuffer())
  if (bytes.byteLength === 0) throw new UniverError(fallbackMessage, 'IMAGE_SOURCE_UNAVAILABLE')
  return bytes
}

function decodeDataUri(source: string): ReferenceImage {
  const match = /^data:([\w.+-]+\/[\w.+-]+);base64,([A-Za-z0-9+/]+={0,2})$/u.exec(source)
  const mediaType = match?.[1]
  const encoded = match?.[2]
  if (mediaType === undefined || encoded === undefined) {
    throw new UniverError(
      'The selected image uses a data URI this feature cannot read.',
      'IMAGE_SOURCE_UNSUPPORTED'
    )
  }
  const bytes = Uint8Array.from(Buffer.from(encoded, 'base64'))
  if (bytes.byteLength === 0) {
    throw new UniverError('The selected image is empty.', 'IMAGE_SOURCE_UNSUPPORTED')
  }
  return { bytes, mediaType }
}

function sniffMediaType(bytes: Uint8Array): string {
  if ([137, 80, 78, 71].every((byte, index) => bytes[index] === byte)) return 'image/png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  if (Buffer.from(bytes.subarray(0, 4)).toString('ascii') === 'RIFF') return 'image/webp'
  const header = Buffer.from(bytes.subarray(0, 6)).toString('ascii')
  if (header === 'GIF87a' || header === 'GIF89a') return 'image/gif'
  throw new UniverError(
    'The selected image is not a PNG, JPEG, WebP, or GIF file.',
    'IMAGE_SOURCE_UNSUPPORTED'
  )
}

function readSourceType(value: unknown): SlideImageElement['sourceType'] {
  return value === 'BASE64' || value === 'UUID' || value === 'URL' ? value : undefined
}

function readDimension(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : 0
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
