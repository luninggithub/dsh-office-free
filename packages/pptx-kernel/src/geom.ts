import { attribute, childElements, findDescendant, type XmlElement } from './xml.ts'

/** Axis-aligned placement accumulated through group transforms. Rotation is applied separately. */
export interface Transform {
  readonly x: number
  readonly y: number
  readonly sx: number
  readonly sy: number
}

export const IDENTITY: Transform = { x: 0, y: 0, sx: 1, sy: 1 }

export interface Xfrm {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly chX: number
  readonly chY: number
  readonly chWidth: number
  readonly chHeight: number
  readonly rotation: number
  readonly flipH: boolean
  readonly flipV: boolean
}

export const EMPTY_XFRM: Xfrm = {
  x: 0,
  y: 0,
  width: 0,
  height: 0,
  chX: 0,
  chY: 0,
  chWidth: 0,
  chHeight: 0,
  rotation: 0,
  flipH: false,
  flipV: false
}

export function readNumber(value: string | undefined, fallback = 0): number {
  if (value === undefined) return fallback
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : fallback
}

/** The shape or group's own transform. An empty `spPr` yields `undefined` so callers can inherit. */
export function shapeXfrm(xml: string, element: XmlElement): Xfrm | undefined {
  const children = childElements(xml, element.openEnd, element.closeStart)
  const props = children.find((item) => item.name === 'spPr' || item.name === 'grpSpPr')
  if (props !== undefined) {
    const nested = findDescendant(xml, 'xfrm', props.openEnd, props.closeStart)
    if (nested !== undefined) return xfrmFrom(xml, nested)
  }
  const own = children.find((item) => item.name === 'xfrm')
  return own === undefined ? undefined : xfrmFrom(xml, own)
}

export function mapBox(
  parent: Transform,
  x: number,
  y: number,
  width: number,
  height: number
): { x: number; y: number; width: number; height: number } {
  return {
    x: parent.x + x * parent.sx,
    y: parent.y + y * parent.sy,
    width: width * parent.sx,
    height: height * parent.sy
  }
}

/**
 * Child coordinates are in the group's `chOff`/`chExt` space.
 * A missing child extent keeps the parent scale instead of dividing by zero.
 */
export function groupTransform(parent: Transform, group: Xfrm): Transform {
  const box = mapBox(parent, group.x, group.y, group.width, group.height)
  const sx = group.chWidth === 0 ? parent.sx : box.width / group.chWidth
  const sy = group.chHeight === 0 ? parent.sy : box.height / group.chHeight
  return {
    x: box.x - group.chX * sx,
    y: box.y - group.chY * sy,
    sx,
    sy
  }
}

/** Clockwise rotation in a y-down coordinate system, matching DrawingML. */
export function rotatePoint(
  x: number,
  y: number,
  pivotX: number,
  pivotY: number,
  degrees: number
): { x: number; y: number } {
  if (degrees === 0) return { x, y }
  const rad = (degrees * Math.PI) / 180
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)
  const dx = x - pivotX
  const dy = y - pivotY
  return {
    x: pivotX + dx * cos - dy * sin,
    y: pivotY + dx * sin + dy * cos
  }
}

function xfrmFrom(xml: string, element: XmlElement): Xfrm {
  const kids = childElements(xml, element.openEnd, element.closeStart)
  const off = kids.find((item) => item.name === 'off')
  const ext = kids.find((item) => item.name === 'ext')
  const chOff = kids.find((item) => item.name === 'chOff')
  const chExt = kids.find((item) => item.name === 'chExt')
  return {
    x: readNumber(off === undefined ? undefined : attribute(off.attrs, 'x')),
    y: readNumber(off === undefined ? undefined : attribute(off.attrs, 'y')),
    width: readNumber(ext === undefined ? undefined : attribute(ext.attrs, 'cx')),
    height: readNumber(ext === undefined ? undefined : attribute(ext.attrs, 'cy')),
    chX: readNumber(chOff === undefined ? undefined : attribute(chOff.attrs, 'x')),
    chY: readNumber(chOff === undefined ? undefined : attribute(chOff.attrs, 'y')),
    chWidth: readNumber(chExt === undefined ? undefined : attribute(chExt.attrs, 'cx')),
    chHeight: readNumber(chExt === undefined ? undefined : attribute(chExt.attrs, 'cy')),
    rotation: readNumber(attribute(element.attrs, 'rot')) / 60000,
    flipH: attribute(element.attrs, 'flipH') === '1',
    flipV: attribute(element.attrs, 'flipV') === '1'
  }
}
