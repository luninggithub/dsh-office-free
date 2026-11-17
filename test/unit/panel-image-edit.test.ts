import { mkdtemp, realpath, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { SessionStore } from '@deepseek-ai/dsh-session'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { panelImageEditRoute } from '../../src/host/webServer/routes/panel.ts'
import type { EditSlideImageRequest } from '../../src/host/service/types.ts'
import type { UniverService } from '../../src/host/service/univer-service.ts'
import type { OpenPanel } from '../../src/shared/wire/panel.ts'

describe('panelImageEditRoute', () => {
  let workspace = ''
  let file = ''
  const scope = { cwd: '' }
  const seen: Array<{ request: EditSlideImageRequest; aborted: boolean }> = []

  beforeAll(async () => {
    workspace = await realpath(await mkdtemp(join(tmpdir(), 'dsh-region-')))
    scope.cwd = workspace
    file = join(workspace, 'deck.univer')
    await writeFile(file, '')
    file = await realpath(file)
  })

  afterAll(async () => {
    if (workspace.length > 0) await rm(workspace, { recursive: true, force: true })
  })

  const sessions = {
    get: (id: string) => (id === 'session' ? { header: { cwd: scope.cwd } } : undefined)
  } as unknown as SessionStore

  function service(panel: OpenPanel | undefined): UniverService {
    return {
      panelOf: () => panel,
      editSlideImage: (request: EditSlideImageRequest, signal?: AbortSignal) => {
        seen.push({ request, aborted: signal?.aborted === true })
        return Promise.resolve({
          ok: true as const,
          operation: 'image-edit' as const,
          file: request.file,
          result: {}
        })
      }
    } as unknown as UniverService
  }

  function openPanel(): OpenPanel {
    return {
      source: null,
      file,
      worktreeId: 'wt-1',
      unitId: 'unit-1',
      viewerUrl: '/univer-viewer/?file=KEY',
      selection: null
    }
  }

  it('rejects a missing panel, a blank prompt, and an unknown session', async () => {
    const body = {
      sessionId: 'session',
      pageId: 'page-1',
      elementId: 'image-1',
      prompt: '修改<point>1 2</point>区域内的内容：亮一点'
    }
    await expect(panelImageEditRoute(service(undefined), sessions, body)).rejects.toMatchObject({
      code: 'IMAGE_SELECTION_MISSING'
    })
    await expect(
      panelImageEditRoute(service(openPanel()), sessions, { ...body, prompt: '  ' })
    ).rejects.toMatchObject({ code: 'INVALID_REQUEST' })
    await expect(
      panelImageEditRoute(service(openPanel()), sessions, { ...body, sessionId: 'other' })
    ).rejects.toMatchObject({ code: 'SESSION_SCOPE_UNAVAILABLE' })
  })

  it('forwards the packaged prompt for the panel file and the abort signal', async () => {
    seen.length = 0
    const controller = new AbortController()
    controller.abort()
    const prompt = '修改<bbox>1 2 3 4</bbox>区域内的内容：换背景'
    await panelImageEditRoute(
      service(openPanel()),
      sessions,
      { sessionId: 'session', pageId: 'page-1', elementId: 'image-1', prompt: `  ${prompt}  ` },
      controller.signal
    )
    expect(seen).toHaveLength(1)
    expect(seen[0]?.aborted).toBe(true)
    expect(seen[0]?.request).toMatchObject({
      file,
      worktreeId: 'wt-1',
      unitId: 'unit-1',
      pageId: 'page-1',
      elementId: 'image-1',
      prompt
    })
  })

  it('refuses a panel file outside the session workspace', async () => {
    const outside = await realpath(await mkdtemp(join(tmpdir(), 'dsh-region-out-')))
    const outsideFile = join(outside, 'other.univer')
    await writeFile(outsideFile, '')
    try {
      await expect(
        panelImageEditRoute(service({ ...openPanel(), file: outsideFile }), sessions, {
          sessionId: 'session',
          pageId: 'page-1',
          elementId: 'image-1',
          prompt: '修改'
        })
      ).rejects.toMatchObject({ code: 'SESSION_SCOPE_DENIED' })
    } finally {
      await rm(outside, { recursive: true, force: true })
    }
  })
})
