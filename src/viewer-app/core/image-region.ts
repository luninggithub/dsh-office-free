/** One marked area, stored as fractions of the image so it scales with the preview. */
export interface ImageRegion {
  readonly id: string
  readonly kind: RegionKind
  /** Left edge, fraction of the image width. */
  readonly x: number
  /** Top edge, fraction of the image height. */
  readonly y: number
  readonly w: number
  readonly h: number
  readonly text: string
}

export type RegionKind = 'point' | 'bbox'

export type Handle = 'nw' | 'ne' | 'sw' | 'se'

export interface ImageSize {
  readonly width: number
  readonly height: number
}

export interface PixelRect {
  readonly left: number
  readonly top: number
  readonly width: number
  readonly height: number
}

export const MAX_REGIONS = 8
export const DEFAULT_REGION_WIDTH = 240
export const DEFAULT_REGION_HEIGHT = 136
export const POINT_REGION_WIDTH = 72
export const POINT_REGION_HEIGHT = 40
/** A shorter movement is a click, which records a point rather than a box. */
export const DRAG_THRESHOLD_PX = 4
const MIN_REGION_PX = 8
const GEOMETRY_EPSILON = 1e-4

/** The box shown when editing starts. It is not itself an edit until the user changes it. */
export function createDefaultRegion(display: ImageSize, id = 'region-1'): ImageRegion {
  const size = normalize(display)
  const width = Math.min(DEFAULT_REGION_WIDTH, size.width) / size.width
  const height = Math.min(DEFAULT_REGION_HEIGHT, size.height) / size.height
  return {
    id,
    kind: 'bbox',
    x: (1 - width) / 2,
    y: (1 - height) / 2,
    w: width,
    h: height,
    text: ''
  }
}

/** A click. The recorded point is the center of the 72×40 box. */
export function createPointRegion(
  display: ImageSize,
  centerX: number,
  centerY: number,
  id: string
): ImageRegion {
  const size = normalize(display)
  const width = Math.min(POINT_REGION_WIDTH, size.width) / size.width
  const height = Math.min(POINT_REGION_HEIGHT, size.height) / size.height
  return clampRegion(
    {
      id,
      kind: 'point',
      x: centerX / size.width - width / 2,
      y: centerY / size.height - height / 2,
      w: width,
      h: height,
      text: ''
    },
    size
  )
}

/** A drag on blank image. The box is the drag rectangle. */
export function createDragRegion(
  display: ImageSize,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  id: string
): ImageRegion {
  const size = normalize(display)
  const left = Math.min(x1, x2)
  const top = Math.min(y1, y2)
  return clampRegion(
    {
      id,
      kind: 'bbox',
      x: left / size.width,
      y: top / size.height,
      w: Math.max(MIN_REGION_PX, Math.abs(x2 - x1)) / size.width,
      h: Math.max(MIN_REGION_PX, Math.abs(y2 - y1)) / size.height,
      text: ''
    },
    size
  )
}

export function clampRegion(region: ImageRegion, display: ImageSize): ImageRegion {
  const size = normalize(display)
  const minW = Math.min(1, MIN_REGION_PX / size.width)
  const minH = Math.min(1, MIN_REGION_PX / size.height)
  const w = clamp(region.w, minW, 1)
  const h = clamp(region.h, minH, 1)
  return {
    ...region,
    w,
    h,
    x: clamp(region.x, 0, 1 - w),
    y: clamp(region.y, 0, 1 - h)
  }
}

/** Moving keeps a point as a point; only the center changes. */
export function moveRegion(
  region: ImageRegion,
  dx: number,
  dy: number,
  display: ImageSize
): ImageRegion {
  const size = normalize(display)
  return clampRegion(
    {
      ...region,
      x: region.x + dx / size.width,
      y: region.y + dy / size.height
    },
    size
  )
}

/** Resizing always defines a box, including when the region started as a click. */
export function resizeRegion(
  region: ImageRegion,
  handle: Handle,
  dx: number,
  dy: number,
  display: ImageSize
): ImageRegion {
  const size = normalize(display)
  const minW = Math.min(1, MIN_REGION_PX / size.width)
  const minH = Math.min(1, MIN_REGION_PX / size.height)
  let x = region.x
  let y = region.y
  let w = region.w
  let h = region.h
  const dfx = dx / size.width
  const dfy = dy / size.height
  if (handle === 'nw' || handle === 'sw') {
    const next = clamp(x + dfx, 0, x + w - minW)
    w = w + (x - next)
    x = next
  } else {
    w = clamp(w + dfx, minW, 1 - x)
  }
  if (handle === 'nw' || handle === 'ne') {
    const next = clamp(y + dfy, 0, y + h - minH)
    h = h + (y - next)
    y = next
  } else {
    h = clamp(h + dfy, minH, 1 - y)
  }
  return clampRegion({ ...region, kind: 'bbox', x, y, w, h }, size)
}

export function isDrag(dx: number, dy: number): boolean {
  return Math.hypot(dx, dy) >= DRAG_THRESHOLD_PX
}

export function canAddRegion(regions: readonly ImageRegion[]): boolean {
  return regions.length < MAX_REGIONS
}

export function nextRegionId(regions: readonly ImageRegion[]): string {
  const used = new Set(regions.map((region) => region.id))
  for (let index = 1; index < 1000; index += 1) {
    const id = `region-${index}`
    if (!used.has(id)) return id
  }
  return `region-${regions.length + 1}`
}

/**
 * True once the user has changed the untouched default box.
 *
 * Typing, adding, deleting, moving, or resizing counts. The initial centered box with an empty
 * input does not, so leaving that state drops it without a confirmation.
 */
export function isDirty(regions: readonly ImageRegion[], baseline: ImageRegion): boolean {
  if (regions.length !== 1) return true
  const only = regions[0]
  if (only === undefined) return true
  return !sameGeometry(only, baseline)
}

export function canApply(regions: readonly ImageRegion[]): boolean {
  return regions.some((region) => region.text.trim().length > 0)
}

export function regionFrame(region: ImageRegion, image: PixelRect): PixelRect {
  return {
    left: image.left + region.x * image.width,
    top: image.top + region.y * image.height,
    width: region.w * image.width,
    height: region.h * image.height
  }
}

/**
 * Map canvas pixels (the engine's CSS pixel space) onto the viewport.
 *
 * `engine` is the size the render engine reports. A canvas stretched by CSS uses a different
 * rectangle, so the ratio converts engine pixels into viewport pixels before adding the canvas
 * origin.
 */
/**
 * Slide objects are keyed `slide-drawing-<unit>-<page>-<drawing id>`, not the drawing id alone.
 * A trailing `-<id>` still names that drawing when the scene cannot look the id up directly.
 */
export function slideObjectKeyMatches(objectKey: string, elementId: string): boolean {
  if (elementId.length === 0) return false
  return objectKey === elementId || objectKey.endsWith(`-${elementId}`)
}

export function clientRectFromCanvas(
  canvas: PixelRect,
  engine: ImageSize,
  topLeft: { readonly x: number; readonly y: number },
  bottomRight: { readonly x: number; readonly y: number }
): PixelRect {
  const scaleX = engine.width > 0 ? canvas.width / engine.width : 1
  const scaleY = engine.height > 0 ? canvas.height / engine.height : 1
  const left = Math.min(topLeft.x, bottomRight.x) * scaleX
  const top = Math.min(topLeft.y, bottomRight.y) * scaleY
  return {
    left: canvas.left + left,
    top: canvas.top + top,
    width: Math.abs(bottomRight.x - topLeft.x) * scaleX,
    height: Math.abs(bottomRight.y - topLeft.y) * scaleY
  }
}

/**
 * Package the regions that have text into the image model's prompt.
 *
 * Coordinates are the bitmap's own pixels, origin at the top-left. A click is the center point.
 * A drag, and the initial 240×136 box, are the top-left then the bottom-right. Empty inputs are
 * omitted. Regions are joined with a full-width semicolon.
 */
export function promptOf(regions: readonly ImageRegion[], bitmap: ImageSize): string {
  if (!(bitmap.width > 0) || !(bitmap.height > 0)) return ''
  const parts: string[] = []
  for (const region of regions) {
    const text = region.text.replace(/\s+/gu, ' ').trim()
    if (text.length === 0) continue
    parts.push(
      region.kind === 'point' ? pointPrompt(region, bitmap, text) : boxPrompt(region, bitmap, text)
    )
  }
  return parts.join('；')
}

function pointPrompt(region: ImageRegion, bitmap: ImageSize, text: string): string {
  const x = pointCoord(region.x + region.w / 2, bitmap.width)
  const y = pointCoord(region.y + region.h / 2, bitmap.height)
  return `修改<point>${x} ${y}</point>区域内的内容：${text}`
}

function boxPrompt(region: ImageRegion, bitmap: ImageSize, text: string): string {
  const [x1, x2] = edge(region.x, region.w, bitmap.width)
  const [y1, y2] = edge(region.y, region.h, bitmap.height)
  return `修改<bbox>${x1} ${y1} ${x2} ${y2}</bbox>区域内的内容：${text}`
}

function edge(start: number, span: number, total: number): [number, number] {
  let a = clamp(Math.round(start * total), 0, total)
  let b = clamp(Math.round((start + span) * total), 0, total)
  if (b < a) {
    const swap = a
    a = b
    b = swap
  }
  if (b === a) {
    if (a >= total) a = Math.max(0, total - 1)
    b = Math.min(total, a + 1)
  }
  return [a, b]
}

function pointCoord(center: number, total: number): number {
  return clamp(Math.round(center * total), 0, total)
}

function sameGeometry(region: ImageRegion, baseline: ImageRegion): boolean {
  return (
    region.kind === baseline.kind &&
    region.text.trim() === baseline.text.trim() &&
    Math.abs(region.x - baseline.x) < GEOMETRY_EPSILON &&
    Math.abs(region.y - baseline.y) < GEOMETRY_EPSILON &&
    Math.abs(region.w - baseline.w) < GEOMETRY_EPSILON &&
    Math.abs(region.h - baseline.h) < GEOMETRY_EPSILON
  )
}

function normalize(display: ImageSize): ImageSize {
  return {
    width: Number.isFinite(display.width) && display.width > 0 ? display.width : 1,
    height: Number.isFinite(display.height) && display.height > 0 ? display.height : 1
  }
}

function clamp(value: number, min: number, max: number): number {
  if (max < min) return min
  return Math.min(max, Math.max(min, value))
}
