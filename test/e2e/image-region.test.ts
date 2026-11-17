// The built Viewer: one selected slide image gets the region editor, and a clean exit drops the
// default box. Generation stays on the Host route tests; this file only drives the overlay.
import type { Browser, Page } from 'puppeteer-core'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { launchRenderBrowser } from '../support/browser.ts'
import { startGateway, type GatewayHarness } from '../support/gateway.ts'

const SOURCE_PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='

describe('slide image region editor', () => {
  let gateway: GatewayHarness
  let browser: Browser
  let worktreeId: string
  let unitId: string
  let pageId: string
  let elementId: string

  beforeAll(async () => {
    gateway = await startGateway()
    worktreeId = await gateway.createWorktree('region editor')
    unitId = await gateway.createUnit(worktreeId, 'slide', 'Deck')
    const inserted = await gateway.execute(
      worktreeId,
      unitId,
      `const slide = presentation.getSlideByIndex(0);
const info = slide
  .newImage()
  .setSource(${JSON.stringify(SOURCE_PNG_DATA_URL)}, api.Enum.ImageSourceType.BASE64)
  .setAbsolutePosition(80, 60)
  .setSize(800, 420)
  .build();
const image = slide.insertImage(info);
if (!image) throw new Error('Cannot insert image');
return { pageId: slide.getId(), elementId: image.getId() };`
    )
    expect(inserted.committed).toBe(true)
    pageId = String(inserted.value.pageId)
    elementId = String(inserted.value.elementId)
    browser = await launchRenderBrowser()
  })

  afterAll(async () => {
    await browser?.close()
    await gateway?.dispose()
  })

  it('puts Edit region 4px under the image and drops an untouched box on Back', async () => {
    const page = await browser.newPage()
    const messages: string[] = []
    page.on('console', (message) => messages.push(`${message.type()}: ${message.text()}`))
    page.on('pageerror', (error) => messages.push(`pageerror: ${String(error)}`))
    page.setDefaultTimeout(30_000)
    try {
      await page.setViewport({ width: 1280, height: 800, deviceScaleFactor: 1 })
      const address = new URL(gateway.origin)
      address.searchParams.set('file', gateway.fileKey)
      address.searchParams.set('mode', 'embedded')
      address.searchParams.set('scope', 'worktree')
      address.searchParams.set('worktree', worktreeId)
      address.searchParams.set('unit', unitId)
      address.searchParams.set('editable', 'true')
      address.searchParams.set('sessionId', 'region-e2e')
      address.searchParams.set('lang', 'zh-CN')
      await page.goto(address.toString(), { waitUntil: 'domcontentloaded' })
      await page.waitForFunction(() => 'univer' in window)
      const selected = await selectImage(page, unitId, pageId, elementId)
      await page
        .waitForSelector('[data-image-region-action="edit"]', { timeout: 25_000 })
        .catch(async (error: unknown) => {
          const debug = await page.evaluate(() => ({
            root: document.querySelector('[data-image-region-root]') !== null,
            canvases: document.querySelectorAll('canvas').length,
            text: document.body.innerText.slice(0, 400)
          }))
          throw new Error(
            `${String(error)} selected=${selected} debug=${JSON.stringify(debug)} messages=${JSON.stringify(messages)}`,
            { cause: error }
          )
        })

      const gap = await page.evaluate(() => {
        const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
        const button = document
          .querySelector('[data-image-region-action="edit"]')
          ?.getBoundingClientRect()
        if (frame === undefined || button === undefined) return null
        return {
          gap: button.top - frame.bottom,
          label: document.querySelector('[data-image-region-action="edit"]')?.textContent,
          width: frame.width,
          height: frame.height
        }
      })
      expect(gap?.label).toBe('编辑区域')
      expect(gap?.width ?? 0).toBeGreaterThan(20)
      expect(gap?.height ?? 0).toBeGreaterThan(20)
      expect(Math.abs((gap?.gap ?? 99) - 4)).toBeLessThan(1.5)

      await page.click('[data-image-region-action="edit"]')
      await page.waitForFunction(() => document.body.innerText.includes('拖拽或点击添加区域'))
      expect(await page.$$('[data-image-region]')).toHaveLength(1)
      await page.click('[data-image-region-action="back"]')
      expect(await page.$('[role="alertdialog"]')).toBeNull()
      await page.waitForSelector('[data-image-region-action="edit"]')

      await page.click('[data-image-region-action="edit"]')
      const corner = await page.evaluate(() => {
        const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
        if (frame === undefined) return null
        return { x: frame.left + 12, y: frame.top + 12 }
      })
      expect(corner).not.toBeNull()
      if (corner === null) return
      await page.mouse.click(corner.x, corner.y)
      await page.waitForSelector('[data-image-region-kind="point"]')
      await page.type('textarea[data-region-id]', '把这里改成蓝色')
      expect(
        await page.$eval(
          '[data-image-region-action="apply"]',
          (element) => (element as HTMLButtonElement).disabled
        )
      ).toBe(false)
      await page.click('[data-image-region-action="back"]')
      await page.waitForSelector('[role="alertdialog"]')
      expect(await page.evaluate(() => document.body.innerText)).toContain(
        '尚未应用的区域和修改内容将丢失。'
      )
      await page.click('[data-image-region-action="confirm-exit"]')
      await page.waitForSelector('[data-image-region-action="edit"]')
      await page.click('[data-image-region-action="edit"]')
      const restored = await page.$eval(
        'textarea',
        (element) => (element as HTMLTextAreaElement).value
      )
      expect(restored).toBe('')
    } finally {
      await page.close()
    }
  })

  it('keeps regions and the edit button on the image while the slide zooms', async () => {
    const page = await browser.newPage()
    page.setDefaultTimeout(30_000)
    try {
      await page.setViewport({ width: 1280, height: 800, deviceScaleFactor: 1 })
      const address = new URL(gateway.origin)
      address.searchParams.set('file', gateway.fileKey)
      address.searchParams.set('mode', 'embedded')
      address.searchParams.set('scope', 'worktree')
      address.searchParams.set('worktree', worktreeId)
      address.searchParams.set('unit', unitId)
      address.searchParams.set('editable', 'true')
      address.searchParams.set('sessionId', 'region-zoom')
      address.searchParams.set('lang', 'zh-CN')
      await page.goto(address.toString(), { waitUntil: 'domcontentloaded' })
      await page.waitForFunction(() => 'univer' in window)
      await selectImage(page, unitId, pageId, elementId)
      await page.waitForSelector('[data-image-region-action="edit"]', { timeout: 25_000 })

      expectAligned(await readAlignment(page))
      const fittedPlacement = await buttonPlacement(page)
      expect(Math.abs(fittedPlacement.gap - 4)).toBeLessThan(1.5)

      const fitted = await frameWidth(page)
      expect(await setZoom(page, unitId, 0.5)).toBe('zoomed')
      await page.waitForFunction(
        (width: number) => {
          const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
          return frame !== undefined && frame.width < width * 0.75
        },
        {},
        fitted
      )
      expectAligned(await readAlignment(page))
      const shrunk = await buttonPlacement(page)
      expect(Math.abs(shrunk.gap - 4)).toBeLessThan(1.5)
      expect(Math.abs(shrunk.centerDelta)).toBeLessThan(2)

      expect(await setZoom(page, unitId, 2)).toBe('zoomed')
      await page.waitForFunction(
        (width: number) => {
          const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
          return frame !== undefined && frame.width > width * 1.6
        },
        {},
        fitted
      )
      const enlarged = await buttonPlacement(page)
      expect(Math.abs(enlarged.gap - 4)).toBeLessThan(1.5)
      expect(enlarged.top).toBeGreaterThan(enlarged.frameTop)

      expect(await setZoom(page, unitId, 0.5)).toBe('zoomed')
      await page.waitForFunction(() => {
        const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
        return frame !== undefined && frame.width < 600
      })
      await page.click('[data-image-region-action="edit"]')
      await page.waitForSelector('[data-image-region]')
      const marked = await readChrome(page)
      expect(marked.fontSize).toBe('13px')

      const markedWidth = await frameWidth(page)
      expect(await setZoom(page, unitId, 0.25)).toBe('zoomed')
      await page.waitForFunction(
        (width: number) => {
          const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
          return frame !== undefined && frame.width < width * 0.7
        },
        {},
        markedWidth
      )
      expectChromeStable(marked, await readChrome(page))

      const beforeWheel = await frameWidth(page)
      await page.evaluate(() => {
        document.querySelector('[data-image-region]')?.dispatchEvent(
          new WheelEvent('wheel', {
            bubbles: true,
            cancelable: true,
            ctrlKey: true,
            deltaY: -160
          })
        )
      })
      await page.waitForFunction(
        (width: number) => {
          const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
          return frame !== undefined && frame.width > width * 1.1
        },
        {},
        beforeWheel
      )
      expectChromeStable(marked, await readChrome(page))
    } finally {
      await page.close()
    }
  })
})

interface Alignment {
  readonly painted: { left: number; top: number; width: number; height: number } | null
  readonly deltaLeft: number | null
  readonly deltaTop: number | null
  readonly deltaWidth: number | null
  readonly deltaHeight: number | null
}

interface Chrome {
  readonly ratioX: number
  readonly ratioY: number
  readonly ratioW: number
  readonly ratioH: number
  readonly fontSize: string
  readonly inputWidth: number
}

function expectAligned(alignment: Alignment): void {
  expect(alignment.painted).not.toBeNull()
  expect(Math.abs(alignment.deltaLeft ?? 99)).toBeLessThan(6)
  expect(Math.abs(alignment.deltaTop ?? 99)).toBeLessThan(6)
  expect(Math.abs(alignment.deltaWidth ?? 99)).toBeLessThan(6)
  expect(Math.abs(alignment.deltaHeight ?? 99)).toBeLessThan(6)
}

function expectChromeStable(before: Chrome, after: Chrome): void {
  expect(Math.abs(after.ratioX - before.ratioX)).toBeLessThan(0.02)
  expect(Math.abs(after.ratioY - before.ratioY)).toBeLessThan(0.02)
  expect(Math.abs(after.ratioW - before.ratioW)).toBeLessThan(0.02)
  expect(Math.abs(after.ratioH - before.ratioH)).toBeLessThan(0.02)
  expect(after.fontSize).toBe('13px')
  expect(Math.abs(after.inputWidth - before.inputWidth)).toBeLessThan(1)
}

async function frameWidth(page: Page): Promise<number> {
  return page.evaluate(() => {
    const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
    return frame?.width ?? 0
  })
}

async function buttonPlacement(page: Page): Promise<{
  gap: number
  centerDelta: number
  top: number
  frameTop: number
}> {
  return page.evaluate(() => {
    const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
    const button = document
      .querySelector('[data-image-region-action="edit"]')
      ?.getBoundingClientRect()
    if (frame === undefined || button === undefined) {
      return { gap: 99, centerDelta: 99, top: 0, frameTop: 0 }
    }
    return {
      gap: button.top - frame.bottom,
      centerDelta: button.left + button.width / 2 - (frame.left + frame.width / 2),
      top: button.top,
      frameTop: frame.top
    }
  })
}

async function readChrome(page: Page): Promise<Chrome> {
  const chrome = await page.evaluate(() => {
    const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
    const region = document.querySelector('[data-image-region]')?.getBoundingClientRect()
    const input = document.querySelector('textarea')
    const box = input?.parentElement
    if (
      frame === undefined ||
      region === undefined ||
      input === null ||
      box === null ||
      box === undefined
    ) {
      return null
    }
    return {
      ratioX: (region.left - frame.left) / frame.width,
      ratioY: (region.top - frame.top) / frame.height,
      ratioW: region.width / frame.width,
      ratioH: region.height / frame.height,
      fontSize: getComputedStyle(input).fontSize,
      inputWidth: box.getBoundingClientRect().width
    }
  })
  expect(chrome).not.toBeNull()
  if (chrome === null) throw new Error('region chrome missing')
  return chrome
}

async function setZoom(page: Page, unitId: string, zoomRatio: number): Promise<string> {
  return page.evaluate(
    (targetUnitId, ratio) => {
      interface InjectorNode {
        children?: InjectorNode[]
        resolvedDependencyCollection?: { resolvedDependencies?: Map<unknown, unknown> }
      }
      const root = (window as unknown as { univer?: { __getInjector(): InjectorNode } }).univer
      if (root === undefined) return 'no-univer'
      const seen = new Set<InjectorNode>()
      const queue = [root.__getInjector()]
      while (queue.length > 0) {
        const injector = queue.shift()
        if (injector === undefined || seen.has(injector)) continue
        seen.add(injector)
        const map = injector.resolvedDependencyCollection?.resolvedDependencies
        if (map instanceof Map) {
          for (const items of map.values()) {
            const list = Array.isArray(items) ? items : [items]
            for (const item of list) {
              const service = item as {
                executeCommand?: (id: string, params: unknown, options: unknown) => Promise<unknown>
              }
              if (typeof service?.executeCommand === 'function') {
                void service.executeCommand(
                  'slide.operation.set-zoom-ratio',
                  { unitId: targetUnitId, zoomRatio: ratio },
                  { onlyLocal: true }
                )
                return 'zoomed'
              }
            }
          }
        }
        queue.push(...(injector.children ?? []))
      }
      return 'missing-command'
    },
    unitId,
    zoomRatio
  )
}

async function readAlignment(page: Page): Promise<Alignment> {
  return page.evaluate(() => {
    const frame = document.querySelector('[data-image-region-frame]')?.getBoundingClientRect()
    const canvas = [...document.querySelectorAll('canvas')].toSorted(
      (left, right) => right.width * right.height - left.width * left.height
    )[0]
    if (frame === undefined || canvas === undefined) {
      return { painted: null, deltaLeft: null, deltaTop: null, deltaWidth: null, deltaHeight: null }
    }
    const context = canvas.getContext('2d', { willReadFrequently: true })
    if (context === null) {
      return { painted: null, deltaLeft: null, deltaTop: null, deltaWidth: null, deltaHeight: null }
    }
    const bounds = canvas.getBoundingClientRect()
    const scaleX = canvas.width / bounds.width
    const scaleY = canvas.height / bounds.height
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
    let minX = canvas.width
    let minY = canvas.height
    let maxX = -1
    let maxY = -1
    for (let y = 0; y < canvas.height; y += 2) {
      for (let x = 0; x < canvas.width; x += 2) {
        const offset = (y * canvas.width + x) * 4
        const red = pixels[offset] ?? 0
        const green = pixels[offset + 1] ?? 0
        const blue = pixels[offset + 2] ?? 0
        if (red < 240 || green > 20 || blue > 20) continue
        minX = Math.min(minX, x)
        minY = Math.min(minY, y)
        maxX = Math.max(maxX, x)
        maxY = Math.max(maxY, y)
      }
    }
    const painted =
      maxX < 0
        ? null
        : {
            left: bounds.left + minX / scaleX,
            top: bounds.top + minY / scaleY,
            width: (maxX - minX) / scaleX,
            height: (maxY - minY) / scaleY
          }
    return {
      painted,
      deltaLeft: painted === null ? null : frame.left - painted.left,
      deltaTop: painted === null ? null : frame.top - painted.top,
      deltaWidth: painted === null ? null : frame.width - painted.width,
      deltaHeight: painted === null ? null : frame.height - painted.height
    }
  })
}

async function selectImage(
  page: Page,
  unitId: string,
  pageId: string,
  elementId: string
): Promise<string> {
  const selected = await page.evaluate(
    (targetUnitId, targetPageId, targetElementId) => {
      interface InjectorNode {
        children?: InjectorNode[]
        resolvedDependencyCollection?: { resolvedDependencies?: Map<unknown, unknown> }
      }
      const root = (window as unknown as { univer?: { __getInjector(): InjectorNode } }).univer
      if (root === undefined) return 'no-univer'
      const seen = new Set<InjectorNode>()
      const queue = [root.__getInjector()]
      while (queue.length > 0) {
        const injector = queue.shift()
        if (injector === undefined || seen.has(injector)) continue
        seen.add(injector)
        const map = injector.resolvedDependencyCollection?.resolvedDependencies
        if (map instanceof Map) {
          for (const items of map.values()) {
            const list = Array.isArray(items) ? items : [items]
            for (const item of list) {
              const service = item as {
                selectDrawings?: (
                  context: { unitId: string; subUnitId: string },
                  ids: string[],
                  focus: string
                ) => void
              }
              if (typeof service?.selectDrawings === 'function') {
                service.selectDrawings(
                  { unitId: targetUnitId, subUnitId: targetPageId },
                  [targetElementId],
                  targetElementId
                )
                return 'selected'
              }
            }
          }
        }
        queue.push(...(injector.children ?? []))
      }
      return 'missing-service'
    },
    unitId,
    pageId,
    elementId
  )
  if (selected === 'selected') return selected
  const center = await page.evaluate(() => {
    const canvas = document.querySelector('canvas')
    if (canvas === null) return null
    const rect = canvas.getBoundingClientRect()
    return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }
  })
  if (center !== null) await page.mouse.click(center.x, center.y)
  return selected
}
