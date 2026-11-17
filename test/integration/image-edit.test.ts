// One Slide image edit end to end: a real Gateway, a real Slide Unit, and a local stand-in for the
// image model, so the reference read, the generated download, and the Facade write are all proven
// without a vendor account or network access.
import { createServer, type Server } from 'node:http'
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'
import { closeServer, listen } from '../support/net.ts'
import { startGateway, type GatewayHarness } from '../support/gateway.ts'

/** Minimal valid 1×1 PNGs. The original and the model result must differ, or a no-op write would pass. */
const SOURCE_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
const RESULT_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const SOURCE_PNG_DATA_URL = `data:image/png;base64,${SOURCE_PNG_BASE64}`
const RESULT_PNG_DATA_URL = `data:image/png;base64,${RESULT_PNG_BASE64}`
const RESULT_PNG_BYTES = Buffer.from(RESULT_PNG_BASE64, 'base64')

describe('slide image edit', () => {
  let harness: GatewayHarness
  let model: Server
  let modelOrigin: string
  const requests: Array<Record<string, unknown>> = []
  const downloads: string[] = []

  beforeAll(async () => {
    model = createServer((request, response) => {
      if (request.method === 'POST' && request.url === '/generate') {
        const chunks: Buffer[] = []
        request.on('data', (chunk: Buffer) => chunks.push(chunk))
        request.on('end', () => {
          requests.push(
            JSON.parse(Buffer.concat(chunks).toString('utf8')) as Record<string, unknown>
          )
          response.writeHead(200, { 'content-type': 'application/json' })
          response.end(JSON.stringify({ data: [{ url: `${modelOrigin}/result.png` }] }))
        })
        return
      }
      if (request.method === 'GET' && request.url === '/result.png') {
        downloads.push(request.url)
        response.writeHead(200, { 'content-type': 'image/png' })
        response.end(RESULT_PNG_BYTES)
        return
      }
      response.writeHead(404)
      response.end()
    })
    modelOrigin = `http://127.0.0.1:${await listen(model)}`
    harness = await startGateway()
  })

  afterAll(async () => {
    await harness?.dispose()
    if (model !== undefined) await closeServer(model)
  })

  it('replaces the selected image with the model result and keeps its frame', async () => {
    const worktreeId = await harness.createWorktree('image edit')
    const unitId = await harness.createUnit(worktreeId, 'slide', 'Deck')
    const inserted = await harness.execute(
      worktreeId,
      unitId,
      `const slide = presentation.getSlideByIndex(0);
const info = slide
  .newImage()
  .setSource(${JSON.stringify(SOURCE_PNG_DATA_URL)}, api.Enum.ImageSourceType.BASE64)
  .setAbsolutePosition(120, 80)
  .setSize(320, 180)
  .build();
const image = slide.insertImage(info);
if (!image) throw new Error('Cannot insert image');
return { pageId: slide.getId(), elementId: image.getId(), source: image.getSource() };`
    )
    expect(inserted.committed).toBe(true)
    const pageId = String(inserted.value.pageId)
    const elementId = String(inserted.value.elementId)

    const service = harness.createService({
      imageEndpoint: `${modelOrigin}/generate`,
      imageApiKey: 'test-key',
      imageModel: 'test-model',
      imageSize: '2K'
    })
    const edited = await service.editSlideImage({
      workspace: harness.scoped.workspace,
      file: harness.scoped.file,
      worktreeId,
      unitId,
      pageId,
      elementId,
      prompt: 'make it a watercolor painting'
    })

    expect(edited).toMatchObject({
      ok: true,
      operation: 'image-edit',
      result: {
        unitId,
        worktreeId,
        pageId,
        elementId,
        prompt: 'make it a watercolor painting',
        model: 'test-model',
        generated: { mediaType: 'image/png', width: 1, height: 1 }
      }
    })

    // The reference reached the model as a data URI of the image the branch actually holds; after
    // the Gateway stores it the element may carry an asset id, so only the URI shape is fixed.
    const payload = requests.at(-1)
    expect(payload?.['model']).toBe('test-model')
    expect(payload?.['prompt']).toBe('make it a watercolor painting')
    expect(payload?.['size']).toBe('2K')
    expect(payload?.['output_format']).toBe('png')
    expect(payload?.['watermark']).toBe(false)
    expect(String(payload?.['image'])).toMatch(/^data:image\/png;base64,[A-Za-z0-9+/=]+$/u)
    expect(downloads).toEqual(['/result.png'])

    const readback = await harness.execute(
      worktreeId,
      unitId,
      `const slide = presentation.getSlideById(${JSON.stringify(pageId)});
const image = slide.getImages().find((candidate) => candidate.getId() === ${JSON.stringify(elementId)});
if (!image) throw new Error('Image element missing');
const transform = image.getTransform();
return { source: image.getSource(), width: transform.width, height: transform.height };`
    )
    expect(readback.value.source).toBe(RESULT_PNG_DATA_URL)
    expect(readback.value).toMatchObject({ width: 320, height: 180 })
  })

  it('rejects a selection that no longer names an image', async () => {
    const worktreeId = await harness.createWorktree('stale selection')
    const unitId = await harness.createUnit(worktreeId, 'slide', 'Deck')
    const pageId = String(
      (
        await harness.execute(
          worktreeId,
          unitId,
          'return { pageId: presentation.getSlideByIndex(0).getId() };'
        )
      ).value.pageId
    )
    const service = harness.createService({
      imageEndpoint: `${modelOrigin}/generate`,
      imageApiKey: 'test-key'
    })
    await expect(
      service.editSlideImage({
        workspace: harness.scoped.workspace,
        file: harness.scoped.file,
        worktreeId,
        unitId,
        pageId,
        elementId: 'missing-element',
        prompt: 'watercolor'
      })
    ).rejects.toMatchObject({ code: 'IMAGE_SELECTION_STALE' })
  })

  it('refuses to edit without an API key', async () => {
    // A developer machine may export a real key; an empty value is the documented "unset" form.
    vi.stubEnv('ARK_API_KEY', '')
    try {
      const service = harness.createService({ imageEndpoint: `${modelOrigin}/generate` })
      await expect(
        service.editSlideImage({
          workspace: harness.scoped.workspace,
          file: harness.scoped.file,
          worktreeId: 'any',
          unitId: 'any',
          pageId: 'any',
          elementId: 'any',
          prompt: 'watercolor'
        })
      ).rejects.toMatchObject({ code: 'IMAGE_API_KEY_MISSING' })
    } finally {
      vi.unstubAllEnvs()
    }
  })
})
