import { UniverError } from '../service/errors.ts'
import type { ImageOutputSize } from '../service/types.ts'

/** Media type of a generated image: the model is asked for PNG, and PNG or JPEG is accepted. */
export type GeneratedImageMediaType = 'image/png' | 'image/jpeg'

/** Bytes of one generated image, with the intrinsic size read from its header. */
export interface GeneratedImage {
  readonly bytes: Uint8Array
  readonly mediaType: GeneratedImageMediaType
  readonly width: number
  readonly height: number
}

/** One edit request against the Volcengine Ark image generation route. */
export interface GenerateImageRequest {
  readonly endpoint: string
  readonly apiKey: string
  readonly model: string
  readonly prompt: string
  /** The selected image, in whatever media type the deck already stores. */
  readonly image: { readonly bytes: Uint8Array; readonly mediaType: string }
  readonly size: ImageOutputSize
  readonly timeoutMs: number
  readonly signal?: AbortSignal
}

/** Largest generated image accepted into a Unit; assets beyond this cannot be stored anyway. */
const MAX_GENERATED_BYTES = 50 * 1024 * 1024

/**
 * Ask the image model to edit the referenced image and download what it returns.
 *
 * Ark answers with a short-lived URL rather than bytes, so the result is fetched here: a
 * model-visible result must not depend on a link the user cannot reproduce. The reference image is
 * sent as a data URI because that is the only form the endpoint accepts inline.
 */
export async function generateImage(request: GenerateImageRequest): Promise<GeneratedImage> {
  const payload = {
    model: request.model,
    prompt: request.prompt,
    image: `data:${request.image.mediaType};base64,${Buffer.from(request.image.bytes).toString('base64')}`,
    size: request.size,
    output_format: 'png',
    watermark: false
  }
  const answer = await postJson(request, payload)
  const url = readResultUrl(answer)
  const bytes = await downloadGeneratedImage(request, url)
  const mediaType = sniffMediaType(bytes)
  const dimensions = readImageSize(bytes, mediaType)
  return { bytes, mediaType, width: dimensions.width, height: dimensions.height }
}

async function postJson(
  request: GenerateImageRequest,
  payload: Readonly<Record<string, unknown>>
): Promise<unknown> {
  const response = await send(request, request.endpoint, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      authorization: `Bearer ${request.apiKey}`
    },
    body: JSON.stringify(payload)
  })
  const text = await response.text()
  if (!response.ok) {
    throw new UniverError(
      `Image model rejected the request: ${describeFailure(text, response.status)}`,
      'IMAGE_GENERATION_FAILED'
    )
  }
  try {
    return JSON.parse(text) as unknown
  } catch (error) {
    throw new UniverError('Image model returned invalid JSON.', 'IMAGE_GENERATION_FAILED', {
      cause: error
    })
  }
}

/** The endpoint may also accept a configured host that pulls results from the same origin. */
function readResultUrl(answer: unknown): string {
  const data = isRecord(answer) ? answer.data : undefined
  const first = Array.isArray(data) ? data[0] : undefined
  const url = isRecord(first) ? first.url : undefined
  if (typeof url !== 'string' || url.length === 0) {
    throw new UniverError('Image model returned no image URL.', 'IMAGE_GENERATION_INVALID_RESPONSE')
  }
  return url
}

async function downloadGeneratedImage(
  request: GenerateImageRequest,
  url: string
): Promise<Uint8Array> {
  let target: URL
  try {
    target = new URL(url)
  } catch (error) {
    throw new UniverError('Image model returned a non-HTTP image URL.', 'IMAGE_GENERATION_FAILED', {
      cause: error
    })
  }
  // Only a plain download: a returned link must not be able to reach a local-only address.
  if (target.protocol !== 'https:' && target.protocol !== 'http:') {
    throw new UniverError('Image model returned a non-HTTP image URL.', 'IMAGE_GENERATION_FAILED')
  }
  const response = await send(request, target.href, { method: 'GET' })
  if (!response.ok) {
    throw new UniverError(
      `Downloading the generated image failed with HTTP ${String(response.status)}.`,
      'IMAGE_GENERATION_FAILED'
    )
  }
  const bytes = new Uint8Array(await response.arrayBuffer())
  if (bytes.byteLength === 0)
    throw new UniverError('The generated image was empty.', 'IMAGE_GENERATION_FAILED')
  if (bytes.byteLength > MAX_GENERATED_BYTES)
    throw new UniverError(
      'The generated image exceeded the 50 MB asset limit.',
      'IMAGE_GENERATION_FAILED'
    )
  return bytes
}

/** Run one request under the operation deadline, reporting timeouts as a stable code. */
async function send(
  request: GenerateImageRequest,
  url: string,
  init: RequestInit
): Promise<Response> {
  const timeout = AbortSignal.timeout(request.timeoutMs)
  const signal = request.signal === undefined ? timeout : AbortSignal.any([request.signal, timeout])
  try {
    return await fetch(url, { ...init, signal })
  } catch (error) {
    if (request.signal?.aborted === true) throw error
    if (isAbortOrTimeout(error)) {
      throw new UniverError(
        `Image generation timed out after ${String(request.timeoutMs)}ms.`,
        'IMAGE_GENERATION_TIMEOUT',
        { cause: error }
      )
    }
    throw new UniverError('Image generation request failed.', 'IMAGE_GENERATION_FAILED', {
      cause: error
    })
  }
}

/** Keep the vendor's message but never echo the API key back into a tool result. */
function describeFailure(text: string, status: number): string {
  try {
    const parsed = JSON.parse(text) as unknown
    const error = isRecord(parsed) ? parsed.error : undefined
    const message = isRecord(error) ? error.message : undefined
    const code = isRecord(error) ? error.code : undefined
    if (typeof message === 'string' && message.length > 0) {
      return typeof code === 'string' && code.length > 0 ? `${code}: ${message}` : message
    }
  } catch {
    // A non-JSON body carries no structured detail; the status is the only stable fact.
  }
  return `HTTP ${String(status)}`
}

function sniffMediaType(bytes: Uint8Array): GeneratedImageMediaType {
  if ([137, 80, 78, 71].every((byte, index) => bytes[index] === byte)) return 'image/png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  throw new UniverError(
    'The generated image is not a PNG or JPEG file.',
    'IMAGE_GENERATION_INVALID_RESPONSE'
  )
}

function readImageSize(
  bytes: Uint8Array,
  mediaType: GeneratedImageMediaType
): { readonly width: number; readonly height: number } {
  if (mediaType === 'image/png') {
    if (bytes.byteLength < 24)
      throw new UniverError('The generated PNG is truncated.', 'IMAGE_GENERATION_INVALID_RESPONSE')
    return { width: readUInt32BE(bytes, 16), height: readUInt32BE(bytes, 20) }
  }
  const jpeg = readJpegSize(bytes)
  if (jpeg === undefined)
    throw new UniverError(
      'The generated JPEG has no readable size.',
      'IMAGE_GENERATION_INVALID_RESPONSE'
    )
  return jpeg
}

/** Walk JPEG segments to the first start-of-frame, which carries the pixel dimensions. */
function readJpegSize(
  bytes: Uint8Array
): { readonly width: number; readonly height: number } | undefined {
  let offset = 2
  while (offset + 9 < bytes.byteLength) {
    if (bytes[offset] !== 0xff) {
      offset += 1
      continue
    }
    const marker = bytes[offset + 1] ?? 0
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      offset += 2
      continue
    }
    const length = ((bytes[offset + 2] ?? 0) << 8) | (bytes[offset + 3] ?? 0)
    if (length < 2) return undefined
    const isFrame =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
    if (isFrame) {
      return {
        height: ((bytes[offset + 5] ?? 0) << 8) | (bytes[offset + 6] ?? 0),
        width: ((bytes[offset + 7] ?? 0) << 8) | (bytes[offset + 8] ?? 0)
      }
    }
    offset += 2 + length
  }
  return undefined
}

function readUInt32BE(bytes: Uint8Array, offset: number): number {
  return (
    (((bytes[offset] ?? 0) << 24) |
      ((bytes[offset + 1] ?? 0) << 16) |
      ((bytes[offset + 2] ?? 0) << 8) |
      (bytes[offset + 3] ?? 0)) >>>
    0
  )
}

function isAbortOrTimeout(error: unknown): boolean {
  let current: unknown = error
  const seen = new Set<object>()
  while (typeof current === 'object' && current !== null && !seen.has(current)) {
    seen.add(current)
    const name = 'name' in current && typeof current.name === 'string' ? current.name : undefined
    const code = 'code' in current && typeof current.code === 'string' ? current.code : undefined
    if (name === 'TimeoutError' || name === 'AbortError' || code === 'ETIMEDOUT') return true
    current = 'cause' in current ? current.cause : undefined
  }
  return false
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
