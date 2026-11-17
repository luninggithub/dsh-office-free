import { describe, expect, it, vi } from 'vitest'
import type { JsonValue } from '../../src/host/service/types.ts'
import type { UniverError } from '../../src/host/service/errors.ts'
import {
  imageDataUrl,
  locateSlideImage,
  readSlideImageBytes,
  replaceSlideImageCode
} from '../../src/host/provider/slide-image.ts'
import { generateImage } from '../../src/host/provider/image-generation-client.ts'

const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1
])

const slideOf = (element: Record<string, unknown>): JsonValue =>
  ({ slides: { page1: { elements: { shape1: element } } } }) as unknown as JsonValue

const imageElement = (overrides: Record<string, unknown> = {}): Record<string, unknown> => ({
  type: 'image',
  source: 'data:image/png;base64,iVBORw0KGgo=',
  imageSourceType: 'BASE64',
  transform: { width: 240, height: 160 },
  ...overrides
})

describe('locateSlideImage', () => {
  it('returns the image with its frame when the reported element is an image', () => {
    expect(
      locateSlideImage(slideOf(imageElement()), { pageId: 'page1', elementId: 'shape1' })
    ).toEqual({
      pageId: 'page1',
      elementId: 'shape1',
      source: 'data:image/png;base64,iVBORw0KGgo=',
      sourceType: 'BASE64',
      width: 240,
      height: 160
    })
  })

  it('rejects a stale page and a stale element', () => {
    const data = slideOf(imageElement())
    expect(() => locateSlideImage(data, { pageId: 'gone', elementId: 'shape1' })).toThrow(
      expect.objectContaining<Partial<UniverError>>({ code: 'IMAGE_SELECTION_STALE' })
    )
    expect(() => locateSlideImage(data, { pageId: 'page1', elementId: 'gone' })).toThrow(
      expect.objectContaining<Partial<UniverError>>({ code: 'IMAGE_SELECTION_STALE' })
    )
  })

  it('rejects an element that is not an image', () => {
    const data = slideOf({ type: 'shape', source: '', transform: { width: 10, height: 10 } })
    expect(() => locateSlideImage(data, { pageId: 'page1', elementId: 'shape1' })).toThrow(
      expect.objectContaining<Partial<UniverError>>({ code: 'IMAGE_SELECTION_NOT_IMAGE' })
    )
  })

  it('rejects an image without a readable source', () => {
    const data = slideOf(imageElement({ source: '' }))
    expect(() => locateSlideImage(data, { pageId: 'page1', elementId: 'shape1' })).toThrow(
      expect.objectContaining<Partial<UniverError>>({ code: 'IMAGE_SOURCE_UNSUPPORTED' })
    )
  })
})

describe('replaceSlideImageCode', () => {
  it('swaps the source of the named image and returns its identity', () => {
    const code = replaceSlideImageCode(
      { pageId: 'page1', elementId: 'shape1' },
      'data:image/png;base64,AAAA'
    )
    expect(code).toContain('presentation.getSlideById("page1")')
    expect(code).toContain('candidate.getId() === "shape1"')
    expect(code).toContain('image.setSource("data:image/png;base64,AAAA", \'BASE64\')')
  })

  it('keeps the frame by replacing content in place', () => {
    const code = replaceSlideImageCode(
      { pageId: 'p', elementId: 'e' },
      'data:image/png;base64,AA=='
    )
    expect(code).not.toContain('setTransform')
    expect(code).not.toContain('insertImage')
  })
})

describe('image bytes', () => {
  it('builds the data URI the Facade stores', () => {
    expect(imageDataUrl({ bytes: PNG, mediaType: 'image/png' })).toBe(
      `data:image/png;base64,${Buffer.from(PNG).toString('base64')}`
    )
  })

  it('decodes a base64 data URI without touching the network', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch')
    const reference = await readSlideImageBytes(
      {
        pageId: 'page1',
        elementId: 'shape1',
        source: imageDataUrl({ bytes: PNG, mediaType: 'image/png' }),
        sourceType: 'BASE64',
        width: 240,
        height: 160
      },
      { gateway: 'http://127.0.0.1:1', fileKey: 'f', worktreeId: 'w', timeoutMs: 1_000 }
    )
    expect(fetchSpy).not.toHaveBeenCalled()
    expect(reference.mediaType).toBe('image/png')
    expect(Buffer.from(reference.bytes)).toEqual(Buffer.from(PNG))
    fetchSpy.mockRestore()
  })

  it('reads a stored asset through the worktree content endpoint', async () => {
    const fetchSpy = vi
      .spyOn(globalThis, 'fetch')
      .mockResolvedValue(new Response(PNG, { status: 200 }))
    const reference = await readSlideImageBytes(
      {
        pageId: 'page1',
        elementId: 'shape1',
        source: 'asset-id',
        sourceType: 'UUID',
        width: 240,
        height: 160
      },
      { gateway: 'http://127.0.0.1:9080/', fileKey: 'file-key', worktreeId: 'wt', timeoutMs: 1_000 }
    )
    expect(fetchSpy.mock.calls[0]?.[0]).toBe(
      'http://127.0.0.1:9080/uf/file-key/worktrees/wt/universer-api/file/asset-id/content'
    )
    expect(reference.mediaType).toBe('image/png')
    fetchSpy.mockRestore()
  })
})

describe('generateImage', () => {
  it('sends the reference as a data URI and returns the downloaded result', async () => {
    const fetchSpy = vi
      .spyOn(globalThis, 'fetch')
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ data: [{ url: 'https://example.test/out.png' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })
      )
      .mockResolvedValueOnce(new Response(PNG, { status: 200 }))

    const generated = await generateImage({
      endpoint: 'https://ark.test/api/v3/images/generations',
      apiKey: 'key',
      model: 'model-1',
      prompt: 'make it watercolor',
      image: { bytes: PNG, mediaType: 'image/jpeg' },
      size: '1K',
      timeoutMs: 1_000
    })

    const [url, init] = fetchSpy.mock.calls[0] ?? []
    expect(url).toBe('https://ark.test/api/v3/images/generations')
    const body = JSON.parse(String((init as RequestInit | undefined)?.body)) as Record<
      string,
      unknown
    >
    expect(body['model']).toBe('model-1')
    expect(body['prompt']).toBe('make it watercolor')
    expect(body['image']).toBe(`data:image/jpeg;base64,${Buffer.from(PNG).toString('base64')}`)
    expect(body['size']).toBe('1K')
    expect(body['output_format']).toBe('png')
    expect(body['watermark']).toBe(false)
    expect(generated.mediaType).toBe('image/png')
    expect(fetchSpy.mock.calls[1]?.[0]).toBe('https://example.test/out.png')
    fetchSpy.mockRestore()
  })

  it('reports a rejected request as a model error', async () => {
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(JSON.stringify({ error: { code: 'ModelNotOpen', message: 'not activated' } }), {
        status: 404,
        headers: { 'content-type': 'application/json' }
      })
    )
    await expect(
      generateImage({
        endpoint: 'https://ark.test/api/v3/images/generations',
        apiKey: 'key',
        model: 'model-1',
        prompt: 'p',
        image: { bytes: PNG, mediaType: 'image/png' },
        size: '1K',
        timeoutMs: 1_000
      })
    ).rejects.toMatchObject({ code: 'IMAGE_GENERATION_FAILED' })
    fetchSpy.mockRestore()
  })

  it('rejects a response without a result URL', async () => {
    const fetchSpy = vi
      .spyOn(globalThis, 'fetch')
      .mockResolvedValue(new Response(JSON.stringify({ data: [] }), { status: 200 }))
    await expect(
      generateImage({
        endpoint: 'https://ark.test/api/v3/images/generations',
        apiKey: 'key',
        model: 'model-1',
        prompt: 'p',
        image: { bytes: PNG, mediaType: 'image/png' },
        size: '1K',
        timeoutMs: 1_000
      })
    ).rejects.toMatchObject({ code: 'IMAGE_GENERATION_INVALID_RESPONSE' })
    fetchSpy.mockRestore()
  })
})
