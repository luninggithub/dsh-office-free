// The built Gateway serves the bundled Viewer, and a Viewer build can appear after that process
// started: a linked development checkout rebuilds `artifacts/viewer` while DSH keeps the same
// Gateway. A startup asset manifest answers 404 for the new hashed chunks even though `index.html`
// on disk already points at them, which leaves the embedded right-panel preview blank.
import { spawn, type ChildProcess } from 'node:child_process'
import { mkdir, mkdtemp, realpath, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { setTimeout as sleep } from 'node:timers/promises'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { freePort } from '../support/net.ts'
import { pluginRoot } from '../support/paths.ts'

const INDEX_HTML =
  '<!doctype html><html><head><script type="module" src="/assets/entry-old.js"></script></head></html>'

/**
 * Start the Gateway artifact directly against a temporary view-assets root.
 *
 * The plugin's own launcher always points the Gateway at the bundled Viewer, so this is the only
 * way to control what the static handler serves; the entry script and readiness line are the same
 * ones the launcher consumes.
 */
async function startGatewayWithAssets(root: string): Promise<{
  readonly origin: string
  readonly child: ChildProcess
}> {
  const port = await freePort()
  const child = spawn(process.execPath, [join(pluginRoot(), 'artifacts', 'gateway.cjs')], {
    env: {
      ...process.env,
      UNIVER_COLLAB_GATEWAY_PORT: String(port),
      UNIVER_VIEW_ASSETS_ROOT: root,
      NODE_PATH: join(pluginRoot(), 'node_modules')
    },
    stdio: ['ignore', 'pipe', 'pipe'],
    windowsHide: true
  })
  await waitForListening(child)
  return { origin: `http://127.0.0.1:${String(port)}`, child }
}

async function waitForListening(child: ChildProcess): Promise<void> {
  const stderr: string[] = []
  child.stderr?.setEncoding('utf8')
  child.stderr?.on('data', (chunk: string) => stderr.push(chunk))
  child.stdout?.setEncoding('utf8')
  const listening = new Promise<void>((resolve, reject) => {
    child.stdout?.on('data', (chunk: string) => {
      if (chunk.includes('listening on')) resolve()
    })
    child.once('exit', (code) => {
      reject(new Error(`Gateway exited before listening (${String(code)}): ${stderr.join('')}`))
    })
  })
  const timeout = sleep(30_000).then(() => {
    throw new Error('Gateway did not report listening within 30s')
  })
  await Promise.race([listening, timeout])
}

describe('built Gateway view assets', () => {
  let root: string
  let gateway: { readonly origin: string; readonly child: ChildProcess }

  beforeAll(async () => {
    root = await realpath(await mkdtemp(join(tmpdir(), 'dsh-univer-view-assets-')))
    await mkdir(join(root, 'assets'))
    await writeFile(join(root, 'index.html'), INDEX_HTML)
    await writeFile(join(root, 'assets', 'entry-old.js'), 'globalThis.viewerBuild = "old"\n')
    gateway = await startGatewayWithAssets(root)
  }, 60_000)

  afterAll(async () => {
    gateway?.child.kill()
    await rm(root, { recursive: true, force: true })
  })

  it('serves a Viewer asset that appeared after the process started', async () => {
    // Replacing the bundle the way a rebuild does: the page keeps its name, the chunk does not.
    await writeFile(join(root, 'index.html'), INDEX_HTML.replace('entry-old.js', 'entry-new.js'))
    await writeFile(join(root, 'assets', 'entry-new.js'), 'globalThis.viewerBuild = "new"\n')

    const chunk = await fetch(`${gateway.origin}/univer-viewer/assets/entry-new.js`)
    expect(chunk.status).toBe(200)
    expect(await chunk.text()).toContain('new')

    // The entry document and the asset it names must come from the same build, or the Viewer
    // loads an empty page instead of the deck.
    const page = await fetch(`${gateway.origin}/univer-viewer/`)
    expect(page.status).toBe(200)
    expect(await page.text()).toContain('entry-new.js')
  })

  it('keeps the previous build serviceable and caches hashed assets', async () => {
    const asset = await fetch(`${gateway.origin}/univer-viewer/assets/entry-old.js`)
    expect(asset.status).toBe(200)
    expect(asset.headers.get('cache-control')).toBe('public, max-age=31536000, immutable')
    const page = await fetch(`${gateway.origin}/univer-viewer/`)
    expect(page.headers.get('cache-control')).toBe('no-cache')
  })
})
