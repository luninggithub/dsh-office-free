import { describe, expect, it } from 'vitest'
import {
  canAddRegion,
  clientRectFromCanvas,
  createDefaultRegion,
  createDragRegion,
  createPointRegion,
  isDirty,
  isDrag,
  moveRegion,
  promptOf,
  resizeRegion,
  slideObjectKeyMatches,
  type ImageRegion
} from '../../src/viewer-app/core/image-region.ts'

const bitmap = { width: 1000, height: 1000 }

describe('image regions', () => {
  it('matches a slide object key that appends the drawing id', () => {
    expect(slideObjectKeyMatches('78pucs', '78pucs')).toBe(true)
    expect(slideObjectKeyMatches('slide-drawing-unit-page-78pucs', '78pucs')).toBe(true)
    expect(slideObjectKeyMatches('slide-drawing-unit-page-78pucs-shape-text', '78pucs')).toBe(false)
    expect(slideObjectKeyMatches('slide-drawing-unit-page-layout:page:0:ph-title', '78pucs')).toBe(
      false
    )
  })

  it('centers a 240×136 box that is not an edit until it changes', () => {
    const display = { width: 800, height: 600 }
    const region = createDefaultRegion(display)
    expect(region.kind).toBe('bbox')
    expect(region.w * display.width).toBeCloseTo(240)
    expect(region.h * display.height).toBeCloseTo(136)
    expect(region.x * display.width).toBeCloseTo((800 - 240) / 2)
    expect(region.y * display.height).toBeCloseTo((600 - 136) / 2)
    expect(isDirty([region], region)).toBe(false)
    expect(isDirty([{ ...region, text: '改颜色' }], region)).toBe(true)
    expect(isDirty([region, region], region)).toBe(true)
  })

  it('keeps a moved click as a point and turns a resize into a box', () => {
    const display = { width: 400, height: 300 }
    const point = createPointRegion(display, 200, 150, 'p')
    expect(point.kind).toBe('point')
    expect(point.w * display.width).toBeCloseTo(72)
    expect(point.h * display.height).toBeCloseTo(40)
    const moved = moveRegion(point, 10, -4, display)
    expect(moved.kind).toBe('point')
    expect((moved.x + moved.w / 2) * display.width).toBeCloseTo(210)
    expect((moved.y + moved.h / 2) * display.height).toBeCloseTo(146)
    const resized = resizeRegion({ ...point, text: 'keep' }, 'se', 20, 10, display)
    expect(resized.kind).toBe('bbox')
    expect(resized.text).toBe('keep')
  })

  it('clamps a click at the edge inside the image', () => {
    const region = createPointRegion({ width: 400, height: 300 }, 1, 1, 'edge')
    expect(region.x).toBeGreaterThanOrEqual(0)
    expect(region.y).toBeGreaterThanOrEqual(0)
    expect(region.x + region.w).toBeLessThanOrEqual(1)
    expect(region.y + region.h).toBeLessThanOrEqual(1)
  })

  it('treats a short movement as a click and stops at eight regions', () => {
    expect(isDrag(3, 0)).toBe(false)
    expect(isDrag(4, 0)).toBe(true)
    const display = { width: 500, height: 400 }
    const drag = createDragRegion(display, 10, 20, 80, 90, 'drag')
    expect(drag.kind).toBe('bbox')
    expect(drag.w * display.width).toBeCloseTo(70)
    const regions = Array.from({ length: 8 }, (_, index) =>
      createDefaultRegion(display, `region-${index + 1}`)
    )
    expect(canAddRegion(regions)).toBe(false)
    expect(canAddRegion(regions.slice(0, 7))).toBe(true)
  })

  it('packages only filled regions in bitmap pixels', () => {
    const box: ImageRegion = {
      id: 'a',
      kind: 'bbox',
      x: 179 / bitmap.width,
      y: 283 / bitmap.height,
      w: (796 - 179) / bitmap.width,
      h: (986 - 283) / bitmap.height,
      text: '输入框1的内容'
    }
    const point: ImageRegion = {
      id: 'b',
      kind: 'point',
      x: (120 - 36) / bitmap.width,
      y: (340 - 20) / bitmap.height,
      w: 72 / bitmap.width,
      h: 40 / bitmap.height,
      text: '  输入框3的内容  '
    }
    const empty: ImageRegion = { ...point, id: 'c', text: '   ' }
    expect(promptOf([box, empty, point], bitmap)).toBe(
      '修改<bbox>179 283 796 986</bbox>区域内的内容：输入框1的内容；修改<point>120 340</point>区域内的内容：输入框3的内容'
    )
    expect(promptOf([empty], bitmap)).toBe('')
    expect(promptOf([box], { width: 0, height: 100 })).toBe('')
  })

  it('scales engine pixels onto the canvas rectangle', () => {
    expect(
      clientRectFromCanvas(
        { left: 10, top: 20, width: 200, height: 100 },
        { width: 100, height: 50 },
        { x: 10, y: 5 },
        { x: 30, y: 15 }
      )
    ).toEqual({ left: 30, top: 30, width: 40, height: 20 })
  })
})
