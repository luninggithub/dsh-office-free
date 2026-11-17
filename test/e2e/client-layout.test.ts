// Source-level browser layout regression for the docked Univer panel: it must sit flush against the
// browser's right edge below the DSH title bar, resize from its left splitter, keep the Viewer
// mounted, and host the review actions itself. The jsdom suite covers the Consumer behaviour; this
// needs real layout and iframe lifecycle.
import { createServer, type Server } from 'node:http'
import { build } from 'esbuild'
import type { Browser, Page } from 'puppeteer-core'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { launchRenderBrowser } from '../support/browser.ts'
import { closeServer, listen } from '../support/net.ts'
import { repoPath } from '../support/paths.ts'

declare global {
  interface Window {
    panelClosed: boolean
    unmountPanel(): void
  }
}

const WORKTREE = 'wt-demo'
/** Title-bar clearances DSH sets on Windows and other desktop shells. */
const CLEARANCES = [0, 40]

describe('docked panel layout in a real browser', () => {
  let server: Server
  let origin: string
  let browser: Browser

  beforeAll(async () => {
    const bundle = await build({
      stdin: {
        contents: `
          import React from 'react'
          import { createRoot } from 'react-dom/client'
          import { UniverPanel } from './src/client/components/univer-panel.tsx'
          import { worktreeStyles } from './src/client/styles/worktree.ts'
          const style = document.createElement('style')
          style.textContent = worktreeStyles
          document.head.append(style)
          const root = createRoot(document.getElementById('root'))
          window.panelClosed = false
          window.unmountPanel = () => root.unmount()
          const unit = {
            unitId: 'u-1', name: 'Sheet', type: 'sheet', kind: 'modified',
            worktreeUrl: '/frame?worktree=${WORKTREE}', mergeUrl: '/frame?worktree=${WORKTREE}&merge=1'
          }
          const state = {
            ok: true, file: 'example.univer', gateway: location.origin,
            gatewayRunning: true, viewerUrl: '/frame?trunk=1',
            worktrees: [{
              worktreeId: '${WORKTREE}', name: 'demo', status: 'draft',
              openUrl: '/frame', worktreeUrl: unit.worktreeUrl, units: [unit]
            }]
          }
          root.render(React.createElement(UniverPanel, {
            file: 'example.univer', state, worktreeId: '${WORKTREE}',
            preferredUnitId: 'u-1', gatewayPhase: 'running',
            startGateway: async () => {}, sessionId: 's1',
            t: key => key, viewerLocale: 'en-US', onState: () => {},
            onClose: () => { window.panelClosed = true; root.unmount() }
          }))
        `,
        resolveDir: repoPath(),
        loader: 'tsx'
      },
      bundle: true,
      write: false,
      platform: 'browser',
      format: 'iife'
    })
    const script = bundle.outputFiles[0]!.text
    server = createServer((request, response) => {
      const url = new URL(request.url ?? '/', 'http://localhost')
      if (url.pathname === '/bundle.js') {
        response.setHeader('Content-Type', 'text/javascript')
        response.end(script)
        return
      }
      if (url.pathname === '/frame') {
        response.setHeader('Content-Type', 'text/html')
        response.end('<button id="focus">Viewer focus target</button>')
        return
      }
      // Every panel API call fails: the test only asserts the failure surface, never state.
      if (url.pathname.startsWith('/univer-api/')) {
        response.writeHead(500, { 'content-type': 'application/json' })
        response.end(JSON.stringify({ ok: false, message: 'probe host' }))
        return
      }
      const clearance = Number(url.searchParams.get('clearance') ?? 0)
      response.setHeader('Content-Type', 'text/html')
      response.end(`<!doctype html><style>
        ${clearance === 0 ? '' : `html{--dsh-frame-top-clearance:${clearance}px}`}
        *{box-sizing:border-box}
        body{margin:0;font:14px system-ui}
        #root{height:100%}
        </style><div id="root"></div><script src="/bundle.js"></script>`)
    })
    origin = `http://127.0.0.1:${await listen(server)}`
    browser = await launchRenderBrowser()
  })

  afterAll(async () => {
    await browser?.close()
    if (server !== undefined) await closeServer(server)
  })

  /** Isolated browser contexts keep the stored panel width from leaking between cases. */
  const contexts = new WeakMap<Page, Awaited<ReturnType<Browser['createBrowserContext']>>>()
  const open = async (query: Record<string, string | number>, width: number): Promise<Page> => {
    const context = await browser.createBrowserContext()
    const page = await context.newPage()
    contexts.set(page, context)
    await page.setViewport({ width, height: 900 })
    const search = new URLSearchParams(
      Object.entries(query).map(([key, value]) => [key, String(value)])
    )
    await page.goto(`${origin}/?${search}`)
    await page.waitForSelector('.uvf_dockFrame')
    return page
  }
  const closePage = async (page: Page): Promise<void> => {
    await page.close()
    await contexts.get(page)?.close()
  }

  const rectOf = (page: Page, selector: string) =>
    page.$eval(selector, (element) => {
      const rect = element.getBoundingClientRect()
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height, bottom: rect.bottom }
    })

  it.each(CLEARANCES)(
    'sits flush against the right edge below a $clearance px title bar',
    async (clearance) => {
      const page = await open({ clearance }, 1280)
      try {
        const dock = await rectOf(page, '.uvf_dock')
        expect(dock).toMatchObject({
          // The default width, and no gap at the browser's right edge.
          width: 560,
          y: clearance + 8,
          bottom: 900
        })
        expect(dock.x + dock.width, 'the dock ends at the viewport edge').toBe(1280)
        expect(await page.evaluate(() => window.scrollY)).toBe(0)
      } finally {
        await closePage(page)
      }
    }
  )

  it('narrows on a small viewport without overflowing it', async () => {
    const page = await open({ clearance: 40 }, 600)
    try {
      const dock = await rectOf(page, '.uvf_dock')
      expect(dock.width, 'no more than 72vw').toBeLessThanOrEqual(432.5)
      expect(dock.x, 'still flush right').toBeGreaterThanOrEqual(0)
      expect(dock.x + dock.width).toBe(600)
      // The narrow layout keeps a 4 px title-bar gap.
      expect(dock.y).toBe(44)
      expect(dock.bottom).toBe(900)
    } finally {
      await closePage(page)
    }
  })

  it('resizes from its left splitter and remembers the width', async () => {
    const page = await open({ clearance: 0 }, 1280)
    try {
      const before = await rectOf(page, '.uvf_dock')
      const handle = (await (await page.$('.uvf_dockSplitter'))!.boundingBox())!
      expect(handle.x, 'the splitter hugs the left border').toBeGreaterThanOrEqual(before.x)
      expect(handle.x).toBeLessThanOrEqual(before.x + 2)
      await page.mouse.move(handle.x + handle.width / 2, 400)
      await page.mouse.down()
      await page.mouse.move(handle.x + handle.width / 2 - 200, 400)
      await page.mouse.up()
      const widened = await rectOf(page, '.uvf_dock')
      expect(widened.width, 'dragging left widens the dock').toBeGreaterThan(before.width)
      expect(widened.x + widened.width, 'the right edge stays flush').toBe(1280)
      const stored = await page.evaluate(() =>
        localStorage.getItem('dsh-univer-office.panel-width')
      )
      expect(stored, 'the width is remembered after the drag').toBe(
        String(Math.round(widened.width))
      )

      // The far edge clamps to the 72vw ceiling instead of eating the page.
      await page.mouse.move(widened.x + 4, 400)
      await page.mouse.down()
      await page.mouse.move(0, 400)
      await page.mouse.up()
      expect(await page.$eval('.uvf_dock', (element) => (element as HTMLElement).style.width)).toBe(
        '922px'
      )
      const clamped = await rectOf(page, '.uvf_dock')

      await page.reload()
      await page.waitForSelector('.uvf_dockFrame')
      expect((await rectOf(page, '.uvf_dock')).width, 'the width survives a reload').toBeCloseTo(
        clamped.width
      )
    } finally {
      await closePage(page)
    }
  })

  it('keeps the Viewer mounted when the title-bar clearance changes', async () => {
    const page = await open({ clearance: 40 }, 1280)
    try {
      const before = await page.evaluate(() => {
        const frame = document.querySelector('iframe')
        ;(frame!.contentWindow as Window & { continuityToken?: string }).continuityToken =
          'preserved'
        return { y: document.querySelector('.uvf_dock')!.getBoundingClientRect().y }
      })
      expect(before.y).toBe(48)
      await page.evaluate(() => {
        document.documentElement.style.setProperty('--dsh-frame-top-clearance', '64px')
        window.dispatchEvent(new Event('resize'))
      })
      await page.waitForFunction(
        () => document.querySelector('.uvf_dock')!.getBoundingClientRect().y === 72
      )
      expect(
        await page.evaluate(
          () =>
            (
              document.querySelector('iframe')!.contentWindow as Window & {
                continuityToken?: string
              }
            ).continuityToken
        ),
        'moving the dock must not reload the Viewer'
      ).toBe('preserved')
    } finally {
      await closePage(page)
    }
  })

  it('hosts the review actions and the confirmation in its own header', async () => {
    const page = await open({ clearance: 0 }, 1280)
    try {
      await page.click('[data-panel-action=compare]')
      await page.waitForSelector('[data-panel-action=compare][aria-pressed=true]')
      const src = await page.$eval('.uvf_dockFrame', (frame) => frame.getAttribute('src') ?? '')
      expect(src, 'the dock asks the embedded Viewer for the diff mode').toContain('view=diff')

      // Opening the confirmation submits nothing; the confirm click below is what reaches the Host,
      // and the simulated Host failure must surface inside the dock.
      await page.click('[data-panel-action=discard]')
      await page.waitForSelector('[role=alertdialog]')
      expect(
        await page.$eval('[role=alertdialog]', (dialog) => dialog.textContent ?? ''),
        'the confirmation states the consequence'
      ).toContain('dock.confirmDiscard')
      await page.click('[data-panel-action=cancel]')
      await page.waitForSelector('[role=alertdialog]', { hidden: true })

      await page.click('[data-panel-action=ready]')
      await page.waitForSelector('.uvf_dockFailure')
      expect(
        await page.$eval('.uvf_dockFailure', (element) => element.textContent ?? '')
      ).toContain('dock.actionFailed')
    } finally {
      await closePage(page)
    }
  })

  it('closes from its own header button', async () => {
    const page = await open({ clearance: 0 }, 1280)
    try {
      await page.click('button[aria-label="dock.close"]')
      await page.waitForFunction(() => window.panelClosed)
      expect(await page.$('.uvf_dock')).toBeNull()
    } finally {
      await closePage(page)
    }
  })
})
