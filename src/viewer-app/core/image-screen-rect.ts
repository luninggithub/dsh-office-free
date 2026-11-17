import type { Univer } from '@univerjs/core'
import type { BaseObject, IRender, Scene } from '@univerjs/engine-render'
import {
  IRenderManagerService,
  ObjectType,
  RENDER_CLASS_TYPE,
  Vector2
} from '@univerjs/engine-render'
import { ISlideDrawingService, PageElementTypeEnum } from '@univerjs-pro/slides'
import { SLIDE_MAIN_VIEWPORT_KEY } from '@univerjs-pro/slides-ui'
import {
  clientRectFromCanvas,
  slideObjectKeyMatches,
  type ImageSize,
  type PixelRect
} from './image-region.ts'

export interface ImageMeasurement {
  readonly rect: PixelRect
  /** Bitmap pixels. Null until the image element has loaded a real size. */
  readonly bitmap: ImageSize | null
}

interface Frame {
  readonly left: number | undefined
  readonly top: number | undefined
  readonly width: number | undefined
  readonly height: number | undefined
}

/**
 * The selected image's viewport rectangle and its bitmap size.
 *
 * Slide pages live either on the main scene or in a page sub-scene drawn through a scene viewer.
 * The viewport converts scene coordinates into canvas pixels; a sub-scene then has the viewer's
 * transform and the main scene's zoom applied, which is the same order the canvas paints them.
 * Returns null until the canvas is mounted and the element can be found.
 */
export function measureSlideImage(
  univer: Univer,
  unitId: string,
  target: { readonly pageId: string; readonly elementId: string }
): ImageMeasurement | null {
  try {
    const injector = univer.__getInjector()
    const render = injector.get(IRenderManagerService).getRenderUnitById(unitId)
    if (render == null) return null
    const drawing = injector.get(ISlideDrawingService).getDrawingByParam({
      unitId,
      subUnitId: target.pageId,
      drawingId: target.elementId
    })
    if (drawing?.element.type !== PageElementTypeEnum.Image) return null
    const transform = drawing.transform
    const frame: Frame | null =
      transform == null
        ? null
        : {
            left: transform.left,
            top: transform.top,
            width: transform.width,
            height: transform.height
          }
    const located = locate(render, target.elementId, frame)
    if (located === null) return null
    const canvas = render.engine.getCanvasElement()
    const canvasRect = canvas.getBoundingClientRect()
    if (canvasRect.width < 2 || canvasRect.height < 2) return null
    if (render.engine.width <= 1 || render.engine.height <= 1) return null
    const bound = located.object.getRealBound()
    const topLeft = sceneToCanvas(render.scene, located.scene, bound.left, bound.top)
    const bottomRight = sceneToCanvas(
      render.scene,
      located.scene,
      bound.left + bound.width,
      bound.top + bound.height
    )
    if (topLeft === null || bottomRight === null) return null
    const rect = clientRectFromCanvas(
      {
        left: canvasRect.left,
        top: canvasRect.top,
        width: canvasRect.width,
        height: canvasRect.height
      },
      { width: render.engine.width, height: render.engine.height },
      topLeft,
      bottomRight
    )
    if (rect.width < 2 || rect.height < 2) return null
    return { rect, bitmap: bitmapOf(located.object) }
  } catch {
    // The slide canvas is not mounted yet, or this Unit has no render scene.
    return null
  }
}

function locate(
  render: IRender,
  elementId: string,
  frame: Frame | null
): { readonly scene: Scene; readonly object: BaseObject } | null {
  const scenes = scenesOf(render)
  for (const scene of scenes) {
    const object = findObject(scene, elementId, frame)
    if (object !== null) return { scene, object }
  }
  return null
}

function scenesOf(render: IRender): Scene[] {
  const scenes: Scene[] = []
  const component = render.mainComponent
  if (component !== null && component !== undefined && isSceneViewer(component)) {
    const active = component.getActiveSubScene()
    if (active != null) scenes.push(active)
    for (const scene of component.getSubScenes().values()) {
      if (!scenes.includes(scene)) scenes.push(scene)
    }
  }
  scenes.push(render.scene)
  return scenes
}

function isSceneViewer(value: object): value is {
  getActiveSubScene(): Scene | null | undefined
  getSubScenes(): Map<string, Scene>
} {
  return 'getActiveSubScene' in value && 'getSubScenes' in value
}

function findObject(scene: Scene, elementId: string, frame: Frame | null): BaseObject | null {
  const direct = scene.getObject(elementId) ?? scene.getObjectIncludeInGroup(elementId)
  if (direct != null) return direct
  let named: BaseObject | null = null
  for (const object of scene.getAllObjects()) {
    if (!slideObjectKeyMatches(object.oKey, elementId)) continue
    if (object.objectType === ObjectType.IMAGE) return object
    named ??= object
  }
  if (named !== null) return named
  if (!hasFrame(frame)) return null
  let best: BaseObject | null = null
  let bestScore = 6
  for (const object of scene.getAllObjects()) {
    if (object.objectType !== ObjectType.IMAGE) continue
    const score =
      Math.abs(object.left - frame.left) +
      Math.abs(object.top - frame.top) +
      Math.abs(object.width - frame.width) +
      Math.abs(object.height - frame.height)
    if (score < bestScore) {
      best = object
      bestScore = score
    }
  }
  return best
}

function hasFrame(frame: Frame | null): frame is {
  left: number
  top: number
  width: number
  height: number
} {
  return (
    frame !== null &&
    frame.left !== undefined &&
    frame.top !== undefined &&
    frame.width !== undefined &&
    frame.height !== undefined
  )
}

function sceneToCanvas(
  main: Scene,
  scene: Scene,
  x: number,
  y: number
): { x: number; y: number } | null {
  const viewport =
    scene.getViewport(SLIDE_MAIN_VIEWPORT_KEY) ??
    scene.getMainViewport() ??
    scene.getDefaultViewport()
  let point = viewport.getAbsoluteVector(new Vector2(x, y))
  const parent = scene.getParent()
  if (parent.classType === RENDER_CLASS_TYPE.SCENE_VIEWER && 'transform' in parent) {
    point = parent.transform.applyPoint(point)
    const mainViewport =
      main.getViewport(SLIDE_MAIN_VIEWPORT_KEY) ??
      main.getMainViewport() ??
      main.getDefaultViewport()
    point = mainViewport.getAbsoluteVector(new Vector2(point.x, point.y))
  }
  return { x: point.x, y: point.y }
}

function bitmapOf(object: BaseObject): ImageSize | null {
  const candidate = object as BaseObject & {
    getNativeSize?: () => { width: number; height: number }
    getNative?: () => { naturalWidth: number; naturalHeight: number } | null | undefined
  }
  if (typeof candidate.getNativeSize === 'function') {
    const size = candidate.getNativeSize()
    if (size.width > 0 && size.height > 0) return { width: size.width, height: size.height }
  }
  if (typeof candidate.getNative === 'function') {
    const native = candidate.getNative()
    if (native != null && native.naturalWidth > 0 && native.naturalHeight > 0) {
      return { width: native.naturalWidth, height: native.naturalHeight }
    }
  }
  return null
}
