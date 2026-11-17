// Bundle compatibility in a real DSH deployment: a temporary Profile loads the published manifest
// and bundle patch through DSH's own compatibility gate, activates the Provider, survives a
// restart, and tears down. Runs only against an independently installed DSH 0.1.7+ package named by
// UNIVER_DSH_RUNTIME_ROOT, and never changes that installation's own profile.
import { copyFile, mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { REPO_ROOT } from '../support/paths.ts'

const runtime = process.env.UNIVER_DSH_RUNTIME_ROOT

describe.skipIf(runtime === undefined)('Host activation in a real DSH Profile', () => {
  let dir: string
  let ctx: any
  let start: () => Promise<any>

  const restart = async () => {
    await ctx.fiber.dispose()
    ctx = await start()
  }

  beforeAll(async () => {
    const req = createRequire(join(runtime!, 'package.json'))
    const { boot, initProfile, loadProfileDirectory, readProfilePatches } = await import(
      pathToFileURL(req.resolve('@deepseek-ai/dsh-app-boot')).href
    )
    dir = await mkdtemp(join(tmpdir(), 'univer-profile-'))
    await mkdir(join(dir, 'node_modules/dsh-univer-office/lib'), { recursive: true })
    await symlink(
      dirname(dirname(req.resolve('@deepseek-ai/cordis/package.json'))),
      join(dir, 'node_modules/@deepseek-ai')
    )
    await Promise.all(
      ['puppeteer-core', '@puppeteer', '@univerjs-pro'].map((dependency) =>
        symlink(join(REPO_ROOT, 'node_modules', dependency), join(dir, 'node_modules', dependency))
      )
    )
    // Keep the published manifest and bundle patch: stripped peers bypass DSH's compatibility
    // gate and can hide a plugin that is skipped in real profiles.
    await Promise.all(
      ['package.json', 'cordis.patch.yml'].map((file) =>
        copyFile(join(REPO_ROOT, file), join(dir, 'node_modules/dsh-univer-office', file))
      )
    )
    await copyFile(
      join(REPO_ROOT, 'lib/index.js'),
      join(dir, 'node_modules/dsh-univer-office/lib/index.js')
    )
    initProfile(dir, ['dsh-univer-office'])
    const profile = loadProfileDirectory('dsh', dir, join(runtime!, 'package.json'))
    expect(profile.layers, 'the real bundle must pass DSH compatibility checks').toHaveLength(1)
    const patchPath = profile.patchPath
    await writeFile(
      patchPath,
      JSON.stringify([
        {
          id: 'univer',
          config: { tools: false, skills: false, telemetry: false, autoStartGateway: false }
        }
      ])
    )
    const root = join(dir, 'cordis.yml')
    await writeFile(root, '[]')
    const profileContext = {
      name: 'probe',
      dir,
      patchPath,
      installAnchor: join(runtime!, 'package.json'),
      // Home and Profile must differ: otherwise the same patch is loaded twice.
      home: join(dir, 'home'),
      cwd: dir,
      startedBundles: ['dsh-univer-office'],
      overlays: [],
      telemetryDisabledEnv: '1'
    }
    start = () =>
      boot('dsh', root, readProfilePatches('dsh', profileContext), (c: any) => {
        c.provide('profileContext', profileContext)
      })
    ctx = await start()
  })

  afterAll(async () => {
    try {
      await ctx?.fiber.dispose()
    } finally {
      if (dir !== undefined) await rm(dir, { recursive: true, force: true })
    }
  })

  it('activates the real Univer Provider', () => {
    expect(ctx.get('univer')).toBeTruthy()
  })

  it('re-activates the same Provider across a restart', async () => {
    await restart()
    expect(ctx.get('univer')).toBeTruthy()
  })
})
