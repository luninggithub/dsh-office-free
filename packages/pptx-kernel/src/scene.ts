import {
  colorElement,
  colorMapFrom,
  cssColor,
  defaultColorMap,
  readTheme,
  rgbOf,
  schemeColor,
  type ColorContext,
  type ColorMap,
  type Theme
} from './color.ts'
import {
  EMPTY_XFRM,
  groupTransform,
  IDENTITY,
  mapBox,
  readNumber,
  rotatePoint,
  shapeXfrm,
  type Transform,
  type Xfrm
} from './geom.ts'
import { readRelations, resolvePart, type Relation } from './rels.ts'
import { attribute, childElements, decodeXml, findDescendant, type XmlElement } from './xml.ts'

/**
 * Visual page for one slide.
 *
 * Draw order matches PowerPoint: background, master decorations, layout decorations,
 * then the slide. Placeholder text inherits list styles and the master text styles.
 * `chOff` / `chExt` scale grouped shapes. A rotated group orbits its children around
 * the group center; a group nested inside another rotated group is only approximate.
 */

export interface PaintStop {
  readonly offset: number
  readonly color: string
}

export interface PaintCrop {
  readonly left: number
  readonly top: number
  readonly right: number
  readonly bottom: number
}

export type PaintFill =
  | { readonly kind: 'none' }
  | { readonly kind: 'solid'; readonly color: string }
  | {
      readonly kind: 'linear'
      readonly angle: number
      readonly scaled: boolean
      readonly stops: readonly PaintStop[]
    }
  | { readonly kind: 'radial'; readonly stops: readonly PaintStop[] }
  | { readonly kind: 'picture'; readonly imagePath: string; readonly crop: PaintCrop | undefined }

export interface PaintLine {
  readonly color: string
  readonly width: number
  readonly dash: 'solid' | 'dash' | 'dot'
}

export interface PaintShadow {
  readonly color: string
  readonly blur: number
  readonly offsetX: number
  readonly offsetY: number
}

export interface PaintRun {
  readonly text: string
  readonly fontSize: number
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
  readonly color: string
  readonly latin: string
  readonly eastAsian: string
}

export interface PaintParagraph {
  readonly align: 'l' | 'ctr' | 'r'
  readonly marginLeft: number
  readonly indent: number
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly linePct: number
  readonly linePt: number
  readonly bullet: string | undefined
  readonly runs: readonly PaintRun[]
}

export interface PaintText {
  readonly anchor: 't' | 'ctr' | 'b'
  readonly padding: {
    readonly left: number
    readonly top: number
    readonly right: number
    readonly bottom: number
  }
  readonly paragraphs: readonly PaintParagraph[]
}

export type PaintGeometry = 'rect' | 'round' | 'ellipse' | 'triangle' | 'diamond' | 'line'

export interface PaintShape {
  readonly id: string
  readonly name: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly rotation: number
  readonly flipH: boolean
  readonly flipV: boolean
  readonly geometry: PaintGeometry
  readonly corner: number
  readonly fill: PaintFill
  readonly line: PaintLine | undefined
  readonly shadow: PaintShadow | undefined
  readonly text: PaintText | undefined
  readonly plain: string
  readonly editable: boolean
  readonly selectable: boolean
}

export interface SlideScene {
  readonly width: number
  readonly height: number
  readonly background: PaintFill
  readonly shapes: readonly PaintShape[]
  readonly imagePaths: readonly string[]
}

type Reader = (name: string) => string | undefined
type Mode = 'master' | 'layout' | 'slide'

interface Span {
  readonly xml: string
  readonly element: XmlElement
}

interface Marker {
  readonly type: string
  readonly idx: string | undefined
  readonly element: XmlElement
}

interface HeaderFooter {
  readonly dt: boolean
  readonly ftr: boolean
  readonly sldNum: boolean
}

interface Deck {
  readonly theme: Theme
  readonly map: ColorMap
  readonly ctx: ColorContext
  readonly images: string[]
  readonly fallback: Span | undefined
  readonly title: Span | undefined
  readonly body: Span | undefined
  readonly other: Span | undefined
  readonly layoutXml: string | undefined
  readonly masterXml: string | undefined
  readonly layoutMarkers: readonly Marker[]
  readonly masterMarkers: readonly Marker[]
  readonly slideKinds: ReadonlySet<string>
  readonly layoutKinds: ReadonlySet<string>
  readonly hf: HeaderFooter
  readonly slideNumber: number
}

interface Walk {
  readonly frame: Transform
  readonly rotation: number
  readonly pivotX: number
  readonly pivotY: number
  readonly groupFill: PaintFill
  readonly mode: Mode
  readonly partName: string
  readonly rels: Map<string, Relation>
}

interface RunDraft {
  sizePt: number
  bold: boolean
  italic: boolean
  underline: boolean
  color: string
  latin: string
  ea: string
}

interface Space {
  readonly kind: 'pct' | 'pt'
  readonly value: number
}

interface ParaDraft {
  align: 'l' | 'ctr' | 'r'
  marginLeft: number
  indent: number
  bullet: string | undefined
  before: Space
  after: Space
  linePct: number
  linePt: number
}

interface FontRef {
  readonly latin: string | undefined
  readonly ea: string | undefined
  readonly color: string | undefined
}

const NO_FILL: PaintFill = { kind: 'none' }
const NO_FONT: FontRef = { latin: undefined, ea: undefined, color: undefined }

export function buildSlideScene(read: Reader, slideIndex: number): SlideScene {
  const presentation = read('ppt/presentation.xml')
  if (presentation === undefined) throw new Error('pptx: presentation.xml is missing')
  const slides = slidePartNames(read, presentation)
  const partName = slides[slideIndex]
  if (partName === undefined) throw new Error(`pptx: slide ${slideIndex} does not exist`)
  const slideXml = read(partName)
  if (slideXml === undefined) throw new Error(`pptx: part ${partName} is missing`)

  const slideRels = partRelations(read, partName)
  const layoutName = targetByType(slideRels, partName, '/slideLayout')
  const layoutXml = layoutName === undefined ? undefined : read(layoutName)
  const layoutRels =
    layoutName === undefined ? new Map<string, Relation>() : partRelations(read, layoutName)
  const masterName =
    layoutName === undefined ? undefined : targetByType(layoutRels, layoutName, '/slideMaster')
  const masterXml = masterName === undefined ? undefined : read(masterName)
  const masterRels =
    masterName === undefined ? new Map<string, Relation>() : partRelations(read, masterName)
  const themeName =
    masterName === undefined ? undefined : targetByType(masterRels, masterName, '/theme')
  const theme = readTheme(themeName === undefined ? undefined : read(themeName))
  const map = effectiveMap(slideXml, layoutXml, masterXml)
  const styles = textStyles(presentation, masterXml)
  const layoutMarkers = layoutXml === undefined ? [] : markersOf(layoutXml)
  const masterMarkers = masterXml === undefined ? [] : markersOf(masterXml)
  const slideMarkers = markersOf(slideXml)
  const images: string[] = []
  const deck: Deck = {
    theme,
    map,
    ctx: { theme, map, placeholder: undefined },
    images,
    fallback: styles.fallback,
    title: styles.title,
    body: styles.body,
    other: styles.other,
    layoutXml,
    masterXml,
    layoutMarkers,
    masterMarkers,
    slideKinds: kindsOf(slideMarkers),
    layoutKinds: kindsOf(layoutMarkers),
    hf: headerFooter(slideXml, layoutXml, masterXml),
    slideNumber: slideIndex + 1
  }
  const size = readSlideSize(presentation)
  const shapes: PaintShape[] = []
  const showLayout = showsMaster(slideXml)
  const showMasterShapes = showLayout && (layoutXml === undefined || showsMaster(layoutXml))
  if (showMasterShapes && masterXml !== undefined && masterName !== undefined) {
    paintPart(deck, masterXml, masterName, masterRels, 'master', shapes)
  }
  if (showLayout && layoutXml !== undefined && layoutName !== undefined) {
    paintPart(deck, layoutXml, layoutName, layoutRels, 'layout', shapes)
  }
  paintPart(deck, slideXml, partName, slideRels, 'slide', shapes)
  return {
    width: size.width,
    height: size.height,
    background: resolveBackground(
      deck,
      slideXml,
      partName,
      slideRels,
      layoutXml,
      layoutName,
      layoutRels,
      masterXml,
      masterName,
      masterRels
    ),
    shapes,
    imagePaths: images
  }
}

function slidePartNames(read: Reader, presentation: string): string[] {
  const relations = readRelations(read('ppt/_rels/presentation.xml.rels') ?? '')
  const list = findDescendant(presentation, 'sldIdLst')
  if (list === undefined) return []
  const names: string[] = []
  for (const slideId of childElements(presentation, list.openEnd, list.closeStart)) {
    if (slideId.name !== 'sldId') continue
    const relationId = attribute(slideId.attrs, 'id', true)
    const target = relationId === undefined ? undefined : relations.get(relationId)
    if (target === undefined || !target.type.endsWith('/slide')) continue
    names.push(resolvePart('ppt/presentation.xml', target.target))
  }
  return names
}

function partRelations(read: Reader, partName: string): Map<string, Relation> {
  const slash = partName.lastIndexOf('/')
  const rels = `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`
  const xml = read(rels)
  return xml === undefined ? new Map() : readRelations(xml)
}

function targetByType(
  rels: Map<string, Relation>,
  partName: string,
  suffix: string
): string | undefined {
  for (const relation of rels.values()) {
    if (relation.type.endsWith(suffix)) return resolvePart(partName, relation.target)
  }
  return undefined
}

function readSlideSize(xml: string): { width: number; height: number } {
  const element = findDescendant(xml, 'sldSz')
  const width = readNumber(element === undefined ? undefined : attribute(element.attrs, 'cx'))
  const height = readNumber(element === undefined ? undefined : attribute(element.attrs, 'cy'))
  return {
    width: width > 0 ? width : 12192000,
    height: height > 0 ? height : 6858000
  }
}

function effectiveMap(
  slideXml: string,
  layoutXml: string | undefined,
  masterXml: string | undefined
): ColorMap {
  const master = masterXml === undefined ? undefined : findDescendant(masterXml, 'clrMap')
  let map = colorMapFrom(master?.attrs) ?? defaultColorMap()
  const layout = layoutXml === undefined ? undefined : overrideMap(layoutXml)
  if (layout !== undefined) map = layout
  const slide = overrideMap(slideXml)
  if (slide !== undefined) map = slide
  return map
}

function overrideMap(xml: string): ColorMap | undefined {
  const host = findDescendant(xml, 'clrMapOvr')
  if (host === undefined) return undefined
  const over = direct(xml, host, 'overrideClrMapping')
  if (over === undefined) return undefined
  return colorMapFrom(over.attrs)
}

function textStyles(
  presentation: string,
  masterXml: string | undefined
): {
  fallback: Span | undefined
  title: Span | undefined
  body: Span | undefined
  other: Span | undefined
} {
  const fallback = findDescendant(presentation, 'defaultTextStyle')
  const tx = masterXml === undefined ? undefined : findDescendant(masterXml, 'txStyles')
  return {
    fallback: fallback === undefined ? undefined : { xml: presentation, element: fallback },
    title:
      tx === undefined || masterXml === undefined
        ? undefined
        : spanChild(masterXml, tx, 'titleStyle'),
    body:
      tx === undefined || masterXml === undefined
        ? undefined
        : spanChild(masterXml, tx, 'bodyStyle'),
    other:
      tx === undefined || masterXml === undefined
        ? undefined
        : spanChild(masterXml, tx, 'otherStyle')
  }
}

function spanChild(xml: string, parent: XmlElement, name: string): Span | undefined {
  const element = direct(xml, parent, name)
  return element === undefined ? undefined : { xml, element }
}

function headerFooter(
  slideXml: string,
  layoutXml: string | undefined,
  masterXml: string | undefined
): HeaderFooter {
  const element =
    findDescendant(slideXml, 'hf') ??
    (layoutXml === undefined ? undefined : findDescendant(layoutXml, 'hf')) ??
    (masterXml === undefined ? undefined : findDescendant(masterXml, 'hf'))
  if (element === undefined) return { dt: true, ftr: true, sldNum: true }
  return {
    dt: attribute(element.attrs, 'dt') !== '0',
    ftr: attribute(element.attrs, 'ftr') !== '0',
    sldNum: attribute(element.attrs, 'sldNum') !== '0'
  }
}

function showsMaster(xml: string): boolean {
  const cSld = findDescendant(xml, 'cSld')
  if (cSld === undefined) return true
  return attribute(cSld.attrs, 'showMasterSp') !== '0'
}

function resolveBackground(
  deck: Deck,
  slideXml: string,
  slideName: string,
  slideRels: Map<string, Relation>,
  layoutXml: string | undefined,
  layoutName: string | undefined,
  layoutRels: Map<string, Relation>,
  masterXml: string | undefined,
  masterName: string | undefined,
  masterRels: Map<string, Relation>
): PaintFill {
  const own = readBackground(slideXml, slideName, slideRels, deck)
  if (own !== undefined) return own
  if (layoutXml !== undefined && layoutName !== undefined) {
    const layout = readBackground(layoutXml, layoutName, layoutRels, deck)
    if (layout !== undefined) return layout
  }
  if (masterXml !== undefined && masterName !== undefined) {
    const master = readBackground(masterXml, masterName, masterRels, deck)
    if (master !== undefined) return master
  }
  return { kind: 'solid', color: '#ffffff' }
}

function readBackground(
  xml: string,
  partName: string,
  rels: Map<string, Relation>,
  deck: Deck
): PaintFill | undefined {
  const cSld = findDescendant(xml, 'cSld')
  if (cSld === undefined) return undefined
  const bg = direct(xml, cSld, 'bg')
  if (bg === undefined) return undefined
  const ref = direct(xml, bg, 'bgRef')
  if (ref !== undefined) return fillFromRef(xml, ref, deck, partName, rels, 'bg')
  const props = direct(xml, bg, 'bgPr')
  if (props === undefined) return undefined
  return containerFill(xml, props, deck.ctx, deck, partName, rels, NO_FILL) ?? NO_FILL
}

function paintPart(
  deck: Deck,
  xml: string,
  partName: string,
  rels: Map<string, Relation>,
  mode: Mode,
  shapes: PaintShape[]
): void {
  const tree = findDescendant(xml, 'spTree')
  if (tree === undefined) return
  paintTree(deck, xml, tree.openEnd, tree.closeStart, shapes, {
    frame: IDENTITY,
    rotation: 0,
    pivotX: 0,
    pivotY: 0,
    groupFill: NO_FILL,
    mode,
    partName,
    rels
  })
}

function paintTree(
  deck: Deck,
  xml: string,
  from: number,
  to: number,
  shapes: PaintShape[],
  walk: Walk
): void {
  for (const element of childElements(xml, from, to)) {
    if (element.name === 'grpSp') {
      if (isHidden(xml, element)) continue
      const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM
      const box = mapBox(walk.frame, xfrm.x, xfrm.y, xfrm.width, xfrm.height)
      const center = rotatePoint(
        box.x + box.width / 2,
        box.y + box.height / 2,
        walk.pivotX,
        walk.pivotY,
        walk.rotation
      )
      const props = direct(xml, element, 'grpSpPr')
      const own =
        props === undefined
          ? undefined
          : containerFill(xml, props, deck.ctx, deck, walk.partName, walk.rels, walk.groupFill)
      paintTree(deck, xml, element.openEnd, element.closeStart, shapes, {
        frame: groupTransform(walk.frame, xfrm),
        rotation: walk.rotation + xfrm.rotation,
        pivotX: walk.rotation === 0 ? center.x : walk.pivotX,
        pivotY: walk.rotation === 0 ? center.y : walk.pivotY,
        groupFill: own === undefined || own.kind === 'none' ? walk.groupFill : own,
        mode: walk.mode,
        partName: walk.partName,
        rels: walk.rels
      })
      continue
    }
    if (
      element.name !== 'sp' &&
      element.name !== 'pic' &&
      element.name !== 'cxnSp' &&
      element.name !== 'graphicFrame'
    ) {
      continue
    }
    if (isHidden(xml, element)) continue
    const marker = markerOf(xml, element)
    if (!includeShape(marker, walk.mode, deck)) continue
    if (element.name === 'graphicFrame') {
      paintTable(deck, xml, element, shapes, walk)
      continue
    }
    const shape = buildShape(deck, xml, element, marker, walk)
    if (shape !== undefined) shapes.push(shape)
  }
}

function includeShape(marker: Marker | undefined, mode: Mode, deck: Deck): boolean {
  if (mode === 'slide' || marker === undefined) return true
  const kind = alias(marker.type)
  if (kind !== 'dt' && kind !== 'ftr' && kind !== 'sldNum') return false
  if (kind === 'dt' && !deck.hf.dt) return false
  if (kind === 'ftr' && !deck.hf.ftr) return false
  if (kind === 'sldNum' && !deck.hf.sldNum) return false
  if (mode === 'layout' && deck.slideKinds.has(kind)) return false
  if (mode === 'master' && (deck.slideKinds.has(kind) || deck.layoutKinds.has(kind))) return false
  return true
}

function buildShape(
  deck: Deck,
  xml: string,
  element: XmlElement,
  marker: Marker | undefined,
  walk: Walk
): PaintShape | undefined {
  const identity = identityOf(xml, element)
  if (identity === undefined) return undefined
  const local = inheritedXfrm(deck, xml, element, marker, walk.mode)
  if (local === undefined) return undefined
  const placed = place(walk, local.x, local.y, local.width, local.height, local.rotation)
  if (placed.width === 0 && placed.height === 0) return undefined
  const visual = visuals(deck, xml, element, walk, walk.groupFill)
  const text = element.name === 'pic' ? undefined : paintText(deck, xml, element, marker, walk.mode)
  return {
    id: walk.mode === 'slide' ? identity.id : `${walk.mode}:${identity.id}`,
    name: identity.name,
    x: placed.x,
    y: placed.y,
    width: placed.width,
    height: placed.height,
    rotation: placed.rotation,
    flipH: local.flipH,
    flipV: local.flipV,
    geometry: visual.geometry,
    corner: visual.corner,
    fill: visual.fill,
    line: visual.line,
    shadow: visual.shadow,
    text,
    plain: plainOf(text),
    editable: walk.mode === 'slide' && text !== undefined,
    selectable: walk.mode === 'slide'
  }
}

function place(
  walk: Walk,
  x: number,
  y: number,
  width: number,
  height: number,
  rotation: number
): { x: number; y: number; width: number; height: number; rotation: number } {
  const placed = mapBox(walk.frame, x, y, width, height)
  const center = rotatePoint(
    placed.x + placed.width / 2,
    placed.y + placed.height / 2,
    walk.pivotX,
    walk.pivotY,
    walk.rotation
  )
  return {
    x: center.x - placed.width / 2,
    y: center.y - placed.height / 2,
    width: placed.width,
    height: placed.height,
    rotation: walk.rotation + rotation
  }
}

function inheritedXfrm(
  deck: Deck,
  xml: string,
  element: XmlElement,
  marker: Marker | undefined,
  mode: Mode
): Xfrm | undefined {
  const own = shapeXfrm(xml, element)
  if (own !== undefined) return own
  const proto = prototypes(deck, marker, mode)
  if (proto.layout !== undefined && deck.layoutXml !== undefined) {
    const layout = shapeXfrm(deck.layoutXml, proto.layout.element)
    if (layout !== undefined) return layout
  }
  if (proto.master !== undefined && deck.masterXml !== undefined) {
    return shapeXfrm(deck.masterXml, proto.master.element)
  }
  return undefined
}

function visuals(
  deck: Deck,
  xml: string,
  element: XmlElement,
  walk: Walk,
  groupFill: PaintFill
): {
  geometry: PaintGeometry
  corner: number
  fill: PaintFill
  line: PaintLine | undefined
  shadow: PaintShadow | undefined
} {
  const spPr = direct(xml, element, 'spPr')
  const style = direct(xml, element, 'style')
  const geometry = geometryOf(deck, xml, element, markerOf(xml, element), walk.mode)
  let fill: PaintFill | undefined
  if (element.name === 'pic') {
    const blip = direct(xml, element, 'blipFill')
    fill =
      blip === undefined
        ? NO_FILL
        : paintFillElement(xml, blip, deck.ctx, deck, walk.partName, walk.rels, groupFill)
  } else if (spPr !== undefined) {
    fill = containerFill(xml, spPr, deck.ctx, deck, walk.partName, walk.rels, groupFill)
  }
  if (fill === undefined && style !== undefined) {
    const ref = direct(xml, style, 'fillRef')
    if (ref !== undefined) fill = fillFromRef(xml, ref, deck, walk.partName, walk.rels, 'fill')
  }
  return {
    geometry: geometry.geometry,
    corner: geometry.corner,
    fill: fill ?? NO_FILL,
    line: shapeLine(xml, spPr, style, deck),
    shadow: shapeShadow(xml, spPr, style, deck)
  }
}

function geometryOf(
  deck: Deck,
  xml: string,
  element: XmlElement,
  marker: Marker | undefined,
  mode: Mode
): { geometry: PaintGeometry; corner: number } {
  const proto = prototypes(deck, marker, mode)
  const spans: Span[] = []
  if (proto.master !== undefined && deck.masterXml !== undefined) {
    spans.push({ xml: deck.masterXml, element: proto.master.element })
  }
  if (proto.layout !== undefined && deck.layoutXml !== undefined) {
    spans.push({ xml: deck.layoutXml, element: proto.layout.element })
  }
  spans.push({ xml, element })
  let geometry: PaintGeometry = 'rect'
  let corner = 0
  for (const span of spans) {
    const spPr = direct(span.xml, span.element, 'spPr')
    const geom = spPr === undefined ? undefined : direct(span.xml, spPr, 'prstGeom')
    if (geom === undefined) continue
    geometry = geometryKind(attribute(geom.attrs, 'prst') ?? 'rect')
    corner = geometry === 'round' ? cornerOf(span.xml, geom) : 0
  }
  return { geometry, corner }
}

function geometryKind(prst: string): PaintGeometry {
  if (prst === 'ellipse' || prst === 'oval') return 'ellipse'
  if (prst === 'triangle' || prst === 'isoscelesTriangle') return 'triangle'
  if (prst === 'diamond') return 'diamond'
  if (prst === 'line' || prst === 'straightConnector1') return 'line'
  if (prst.includes('round') && prst.includes('Rect')) return 'round'
  return 'rect'
}

function cornerOf(xml: string, geom: XmlElement): number {
  const list = direct(xml, geom, 'avLst')
  const guides = list === undefined ? [] : childElements(xml, list.openEnd, list.closeStart)
  const adj = guides.find((item) => attribute(item.attrs, 'name') === 'adj')
  const formula = adj === undefined ? '' : (attribute(adj.attrs, 'fmla') ?? '')
  const match = /^val\s+(-?\d+(?:\.\d+)?)/.exec(formula)
  if (match?.[1] !== undefined) return readNumber(match[1]) / 100000
  return 16667 / 100000
}

function paintTable(
  deck: Deck,
  xml: string,
  element: XmlElement,
  shapes: PaintShape[],
  walk: Walk
): void {
  const table = findDescendant(xml, 'tbl', element.openEnd, element.closeStart)
  if (table === undefined) return
  const identity = identityOf(xml, element)
  const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM
  const grid = direct(xml, table, 'tblGrid')
  const cols =
    grid === undefined
      ? []
      : childElements(xml, grid.openEnd, grid.closeStart)
          .filter((item) => item.name === 'gridCol')
          .map((item) => readNumber(attribute(item.attrs, 'w')))
  const rows = childElements(xml, table.openEnd, table.closeStart).filter(
    (item) => item.name === 'tr'
  )
  let top = xfrm.y
  rows.forEach((row, rowIndex) => {
    const height = readNumber(attribute(row.attrs, 'h'))
    let left = xfrm.x
    let col = 0
    const cells = childElements(xml, row.openEnd, row.closeStart).filter(
      (item) => item.name === 'tc'
    )
    for (const cell of cells) {
      if (attribute(cell.attrs, 'hMerge') === '1') continue
      const span = Math.max(1, readNumber(attribute(cell.attrs, 'gridSpan'), 1))
      const width = sumCols(cols, col, span)
      if (attribute(cell.attrs, 'vMerge') !== '1' && width > 0 && height > 0) {
        const placed = place(walk, left, top, width, height, xfrm.rotation)
        const text = cellText(deck, xml, cell)
        const props = direct(xml, cell, 'tcPr')
        const fill =
          props === undefined
            ? NO_FILL
            : (containerFill(xml, props, deck.ctx, deck, walk.partName, walk.rels, NO_FILL) ??
              NO_FILL)
        shapes.push({
          id: `${identity?.id ?? 'table'}:${rowIndex}:${col}`,
          name: identity?.name ?? '',
          x: placed.x,
          y: placed.y,
          width: placed.width,
          height: placed.height,
          rotation: placed.rotation,
          flipH: false,
          flipV: false,
          geometry: 'rect',
          corner: 0,
          fill,
          line: props === undefined ? undefined : cellLine(xml, props, deck),
          shadow: undefined,
          text,
          plain: plainOf(text),
          editable: false,
          selectable: false
        })
      }
      left += width
      col += span
    }
    top += height
  })
}

function sumCols(cols: readonly number[], start: number, span: number): number {
  let total = 0
  for (let index = 0; index < span; index += 1) total += cols[start + index] ?? 0
  return total
}

function cellLine(xml: string, props: XmlElement, deck: Deck): PaintLine | undefined {
  const edge = ['lnL', 'lnR', 'lnT', 'lnB']
    .map((name) => direct(xml, props, name))
    .find((item) => item !== undefined)
  if (edge === undefined) return undefined
  return lineFrom(xml, edge, deck.ctx)
}

function cellText(deck: Deck, xml: string, cell: XmlElement): PaintText | undefined {
  const body = direct(xml, cell, 'txBody')
  if (body === undefined) return undefined
  const own = direct(xml, body, 'lstStyle')
  const rest = own === undefined ? [] : [{ xml, element: own }]
  const props = direct(xml, cell, 'tcPr')
  return paragraphs(deck, xml, body, deck.fallback, NO_FONT, rest, paddingOf(props), 1)
}

function paddingOf(props: XmlElement | undefined): PaintText['padding'] {
  return {
    left: readNumber(props === undefined ? undefined : attribute(props.attrs, 'marL'), 91440),
    top: readNumber(props === undefined ? undefined : attribute(props.attrs, 'marT'), 45720),
    right: readNumber(props === undefined ? undefined : attribute(props.attrs, 'marR'), 91440),
    bottom: readNumber(props === undefined ? undefined : attribute(props.attrs, 'marB'), 45720)
  }
}

function paintText(
  deck: Deck,
  xml: string,
  element: XmlElement,
  marker: Marker | undefined,
  mode: Mode
): PaintText | undefined {
  const body = direct(xml, element, 'txBody')
  if (body === undefined) return undefined
  const metrics = bodyMetrics(deck, xml, element, marker, mode)
  return paragraphs(
    deck,
    xml,
    body,
    deck.fallback,
    fontRefOf(xml, element, deck.ctx),
    styleSpans(deck, xml, element, marker, mode),
    metrics.padding,
    metrics.fontScale,
    metrics.anchor
  )
}

function paragraphs(
  deck: Deck,
  xml: string,
  body: XmlElement,
  fallback: Span | undefined,
  font: FontRef,
  rest: readonly Span[],
  padding: PaintText['padding'],
  fontScale: number,
  anchor: PaintText['anchor'] = 't'
): PaintText {
  const built: PaintParagraph[] = []
  const counters = new Map<number, number>()
  for (const paragraph of childElements(xml, body.openEnd, body.closeStart)) {
    if (paragraph.name !== 'p') continue
    const props = direct(xml, paragraph, 'pPr')
    const level = props === undefined ? 0 : readNumber(attribute(props.attrs, 'lvl'))
    const drafted = draftFor(deck, level, fallback, font, rest)
    if (props !== undefined) applyParaProps(xml, props, drafted.run, drafted.para, deck.ctx)
    const runs = runsOf(xml, paragraph, drafted.run, deck, fontScale)
    if (runs.length === 0) {
      const end = direct(xml, paragraph, 'endParaRPr')
      const next = { ...drafted.run }
      if (end !== undefined) applyRun(xml, end, next, deck.ctx)
      runs.push(toRun(next, '', deck.theme, fontScale))
    }
    const first = runs[0]
    const fontPt = first === undefined ? drafted.run.sizePt * fontScale : first.fontSize
    built.push({
      align: drafted.para.align,
      marginLeft: drafted.para.marginLeft,
      indent: drafted.para.indent,
      spaceBefore: spacePt(drafted.para.before, fontPt),
      spaceAfter: spacePt(drafted.para.after, fontPt),
      linePct: drafted.para.linePct,
      linePt: drafted.para.linePt,
      bullet: bulletOf(drafted.para.bullet, level, counters),
      runs
    })
  }
  return { anchor, padding, paragraphs: built }
}

function styleSpans(
  deck: Deck,
  xml: string,
  element: XmlElement,
  marker: Marker | undefined,
  mode: Mode
): Span[] {
  const spans: Span[] = []
  const bucket = bucketSpan(deck, marker?.type)
  if (bucket !== undefined) spans.push(bucket)
  const proto = prototypes(deck, marker, mode)
  if (proto.master !== undefined && deck.masterXml !== undefined) {
    const list = listStyle(deck.masterXml, proto.master.element)
    if (list !== undefined) spans.push({ xml: deck.masterXml, element: list })
  }
  if (proto.layout !== undefined && deck.layoutXml !== undefined) {
    const list = listStyle(deck.layoutXml, proto.layout.element)
    if (list !== undefined) spans.push({ xml: deck.layoutXml, element: list })
  }
  const own = listStyle(xml, element)
  if (own !== undefined) spans.push({ xml, element: own })
  return spans
}

function bucketSpan(deck: Deck, type: string | undefined): Span | undefined {
  if (type === undefined) return undefined
  const kind = alias(type)
  if (kind === 'title') return deck.title
  if (kind === 'body') return deck.body
  return deck.other
}

function listStyle(xml: string, element: XmlElement): XmlElement | undefined {
  const body = direct(xml, element, 'txBody')
  if (body === undefined) return undefined
  return direct(xml, body, 'lstStyle')
}

function bodyMetrics(
  deck: Deck,
  xml: string,
  element: XmlElement,
  marker: Marker | undefined,
  mode: Mode
): { anchor: PaintText['anchor']; padding: PaintText['padding']; fontScale: number } {
  let anchor: PaintText['anchor'] = 't'
  let left = 91440
  let top = 45720
  let right = 91440
  let bottom = 45720
  let fontScale = 1
  const proto = prototypes(deck, marker, mode)
  const spans: Span[] = []
  if (proto.master !== undefined && deck.masterXml !== undefined) {
    spans.push({ xml: deck.masterXml, element: proto.master.element })
  }
  if (proto.layout !== undefined && deck.layoutXml !== undefined) {
    spans.push({ xml: deck.layoutXml, element: proto.layout.element })
  }
  spans.push({ xml, element })
  for (const span of spans) {
    const body = direct(span.xml, span.element, 'txBody')
    const props = body === undefined ? undefined : direct(span.xml, body, 'bodyPr')
    if (props === undefined) continue
    const value = attribute(props.attrs, 'anchor')
    if (value === 'ctr' || value === 'b' || value === 't') anchor = value
    const lIns = attribute(props.attrs, 'lIns')
    if (lIns !== undefined) left = readNumber(lIns)
    const tIns = attribute(props.attrs, 'tIns')
    if (tIns !== undefined) top = readNumber(tIns)
    const rIns = attribute(props.attrs, 'rIns')
    if (rIns !== undefined) right = readNumber(rIns)
    const bIns = attribute(props.attrs, 'bIns')
    if (bIns !== undefined) bottom = readNumber(bIns)
    const fit = direct(span.xml, props, 'normAutofit')
    if (fit !== undefined) {
      const scale = attribute(fit.attrs, 'fontScale')
      fontScale = scale === undefined ? 1 : readNumber(scale) / 100000
    }
  }
  return { anchor, padding: { left, top, right, bottom }, fontScale }
}

function prototypes(
  deck: Deck,
  marker: Marker | undefined,
  mode: Mode
): { layout: Marker | undefined; master: Marker | undefined } {
  if (marker === undefined || mode === 'master') return { layout: undefined, master: undefined }
  if (mode === 'layout') {
    return {
      layout: undefined,
      master: deck.masterXml === undefined ? undefined : findPrototype(deck.masterMarkers, marker)
    }
  }
  const layout =
    deck.layoutXml === undefined ? undefined : findPrototype(deck.layoutMarkers, marker)
  const masterKey = layout ?? marker
  return {
    layout,
    master: deck.masterXml === undefined ? undefined : findPrototype(deck.masterMarkers, masterKey)
  }
}

function draftFor(
  deck: Deck,
  level: number,
  fallback: Span | undefined,
  font: FontRef,
  rest: readonly Span[]
): { run: RunDraft; para: ParaDraft } {
  const run: RunDraft = {
    sizePt: 18,
    bold: false,
    italic: false,
    underline: false,
    color: cssColor(schemeColor('tx1', deck.ctx)),
    latin: '+mn-lt',
    ea: '+mn-ea'
  }
  const para = emptyPara()
  if (fallback !== undefined) applyLevel(fallback.xml, fallback.element, level, run, para, deck.ctx)
  if (font.latin !== undefined) run.latin = font.latin
  if (font.ea !== undefined) run.ea = font.ea
  if (font.color !== undefined) run.color = font.color
  for (const span of rest) applyLevel(span.xml, span.element, level, run, para, deck.ctx)
  return { run, para }
}

function emptyPara(): ParaDraft {
  return {
    align: 'l',
    marginLeft: 0,
    indent: 0,
    bullet: undefined,
    before: { kind: 'pt', value: 0 },
    after: { kind: 'pt', value: 0 },
    linePct: 1,
    linePt: 0
  }
}

function applyLevel(
  xml: string,
  container: XmlElement,
  level: number,
  run: RunDraft,
  para: ParaDraft,
  ctx: ColorContext
): void {
  const element = levelElement(xml, container, level)
  if (element === undefined) return
  applyParaProps(xml, element, run, para, ctx)
}

function levelElement(xml: string, container: XmlElement, level: number): XmlElement | undefined {
  const children = childElements(xml, container.openEnd, container.closeStart)
  for (let index = level; index >= 0; index -= 1) {
    const found = children.find((item) => item.name === `lvl${index + 1}pPr`)
    if (found !== undefined) return found
  }
  return children.find((item) => item.name === 'defPPr')
}

function applyParaProps(
  xml: string,
  element: XmlElement,
  run: RunDraft,
  para: ParaDraft,
  ctx: ColorContext
): void {
  const align = attribute(element.attrs, 'algn')
  if (align === 'l' || align === 'ctr' || align === 'r') para.align = align
  if (align === 'just') para.align = 'l'
  const margin = attribute(element.attrs, 'marL')
  if (margin !== undefined) para.marginLeft = readNumber(margin)
  const indent = attribute(element.attrs, 'indent')
  if (indent !== undefined) para.indent = readNumber(indent)
  const kids = childElements(xml, element.openEnd, element.closeStart)
  const bullet = kids.find(
    (item) => item.name === 'buNone' || item.name === 'buChar' || item.name === 'buAutoNum'
  )
  if (bullet?.name === 'buNone') para.bullet = undefined
  if (bullet?.name === 'buChar') para.bullet = attribute(bullet.attrs, 'char') ?? '•'
  if (bullet?.name === 'buAutoNum') para.bullet = 'auto'
  const before = kids.find((item) => item.name === 'spcBef')
  if (before !== undefined) para.before = readSpace(xml, before)
  const after = kids.find((item) => item.name === 'spcAft')
  if (after !== undefined) para.after = readSpace(xml, after)
  const line = kids.find((item) => item.name === 'lnSpc')
  if (line !== undefined) {
    const spacing = readLine(xml, line)
    para.linePct = spacing.linePct
    para.linePt = spacing.linePt
  }
  const props = kids.find((item) => item.name === 'defRPr')
  if (props !== undefined) applyRun(xml, props, run, ctx)
}

function readSpace(xml: string, element: XmlElement): Space {
  const pct = direct(xml, element, 'spcPct')
  if (pct !== undefined)
    return { kind: 'pct', value: readNumber(attribute(pct.attrs, 'val')) / 100000 }
  const points = direct(xml, element, 'spcPts')
  if (points !== undefined) {
    return { kind: 'pt', value: readNumber(attribute(points.attrs, 'val')) / 100 }
  }
  return { kind: 'pt', value: 0 }
}

function readLine(xml: string, element: XmlElement): { linePct: number; linePt: number } {
  const pct = direct(xml, element, 'spcPct')
  if (pct !== undefined) {
    return { linePct: readNumber(attribute(pct.attrs, 'val')) / 100000, linePt: 0 }
  }
  const points = direct(xml, element, 'spcPts')
  if (points !== undefined) {
    return { linePct: 1, linePt: readNumber(attribute(points.attrs, 'val')) / 100 }
  }
  return { linePct: 1, linePt: 0 }
}

function spacePt(space: Space, fontPt: number): number {
  return space.kind === 'pt' ? space.value : space.value * fontPt
}

function runsOf(
  xml: string,
  paragraph: XmlElement,
  base: RunDraft,
  deck: Deck,
  fontScale: number
): PaintRun[] {
  const runs: PaintRun[] = []
  for (const child of childElements(xml, paragraph.openEnd, paragraph.closeStart)) {
    if (child.name !== 'r' && child.name !== 'fld' && child.name !== 'br') continue
    const next = { ...base }
    const props = direct(xml, child, 'rPr')
    if (props !== undefined) applyRun(xml, props, next, deck.ctx)
    if (child.name === 'br') {
      runs.push(toRun(next, '\n', deck.theme, fontScale))
      continue
    }
    const text =
      child.name === 'fld' && attribute(child.attrs, 'type') === 'slidenum'
        ? String(deck.slideNumber)
        : inlineText(xml, child)
    runs.push(toRun(next, text, deck.theme, fontScale))
  }
  return runs
}

function applyRun(xml: string, props: XmlElement, run: RunDraft, ctx: ColorContext): void {
  const size = attribute(props.attrs, 'sz')
  if (size !== undefined) run.sizePt = readNumber(size) / 100
  const bold = attribute(props.attrs, 'b')
  if (bold === '1' || bold === 'true') run.bold = true
  if (bold === '0' || bold === 'false') run.bold = false
  const italic = attribute(props.attrs, 'i')
  if (italic === '1' || italic === 'true') run.italic = true
  if (italic === '0' || italic === 'false') run.italic = false
  const underline = attribute(props.attrs, 'u')
  if (underline !== undefined) run.underline = underline !== 'none'
  const kids = childElements(xml, props.openEnd, props.closeStart)
  const solid = kids.find((item) => item.name === 'solidFill')
  if (solid !== undefined) {
    const color = colorElement(xml, solid)
    if (color !== undefined) run.color = cssColor(rgbOf(xml, color, ctx))
  }
  const latin = kids.find((item) => item.name === 'latin')
  const latinFace = latin === undefined ? undefined : attribute(latin.attrs, 'typeface')
  if (latinFace !== undefined && latinFace !== '') run.latin = latinFace
  const ea = kids.find((item) => item.name === 'ea')
  const eaFace = ea === undefined ? undefined : attribute(ea.attrs, 'typeface')
  if (eaFace !== undefined && eaFace !== '') run.ea = eaFace
}

function toRun(run: RunDraft, text: string, theme: Theme, fontScale: number): PaintRun {
  return {
    text,
    fontSize: run.sizePt * fontScale,
    bold: run.bold,
    italic: run.italic,
    underline: run.underline,
    color: run.color,
    latin: resolveFace(run.latin, theme),
    eastAsian: resolveFace(run.ea, theme)
  }
}

function resolveFace(token: string, theme: Theme): string {
  if (token === '+mj-lt') return theme.major.latin || 'Calibri'
  if (token === '+mn-lt') return theme.minor.latin || 'Calibri'
  if (token === '+mj-ea') return theme.major.ea || theme.major.hans || theme.major.latin || '宋体'
  if (token === '+mn-ea') return theme.minor.ea || theme.minor.hans || theme.minor.latin || '宋体'
  if (token === '+mj-cs') return theme.major.cs || theme.major.latin || 'Calibri'
  if (token === '+mn-cs') return theme.minor.cs || theme.minor.latin || 'Calibri'
  return token
}

function fontRefOf(xml: string, element: XmlElement, ctx: ColorContext): FontRef {
  const style = direct(xml, element, 'style')
  const ref = style === undefined ? undefined : direct(xml, style, 'fontRef')
  if (ref === undefined) return NO_FONT
  const token = attribute(ref.attrs, 'idx') === 'major' ? '+mj' : '+mn'
  const color = colorElement(xml, ref)
  return {
    latin: `${token}-lt`,
    ea: `${token}-ea`,
    color: color === undefined ? undefined : cssColor(rgbOf(xml, color, ctx))
  }
}

function bulletOf(
  bullet: string | undefined,
  level: number,
  counters: Map<number, number>
): string | undefined {
  if (bullet !== 'auto') {
    counters.set(level, 0)
    return bullet
  }
  const next = (counters.get(level) ?? 0) + 1
  counters.set(level, next)
  const deeper: number[] = []
  for (const key of counters.keys()) {
    if (key > level) deeper.push(key)
  }
  for (const key of deeper) counters.delete(key)
  return `${next}.`
}

function plainOf(text: PaintText | undefined): string {
  if (text === undefined) return ''
  return text.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n')
}

function inlineText(xml: string, parent: XmlElement): string {
  return childElements(xml, parent.openEnd, parent.closeStart)
    .filter((item) => item.name === 't')
    .map((item) => decodeXml(xml.slice(item.openEnd, item.closeStart)))
    .join('')
}

function containerFill(
  xml: string,
  container: XmlElement,
  ctx: ColorContext,
  deck: Deck,
  partName: string,
  rels: Map<string, Relation>,
  groupFill: PaintFill
): PaintFill | undefined {
  const fill = childElements(xml, container.openEnd, container.closeStart).find((item) =>
    ['noFill', 'solidFill', 'gradFill', 'blipFill', 'grpFill'].includes(item.name)
  )
  if (fill === undefined) return undefined
  return paintFillElement(xml, fill, ctx, deck, partName, rels, groupFill)
}

function paintFillElement(
  xml: string,
  element: XmlElement,
  ctx: ColorContext,
  deck: Deck,
  partName: string,
  rels: Map<string, Relation>,
  groupFill: PaintFill
): PaintFill {
  if (element.name === 'noFill' || element.name === 'grpFill') {
    return element.name === 'grpFill' ? groupFill : NO_FILL
  }
  if (element.name === 'solidFill') {
    const color = colorElement(xml, element)
    return {
      kind: 'solid',
      color: color === undefined ? '#000000' : cssColor(rgbOf(xml, color, ctx))
    }
  }
  if (element.name === 'gradFill') return gradientFill(xml, element, ctx)
  if (element.name === 'blipFill') return pictureFill(xml, element, deck, partName, rels)
  return NO_FILL
}

function fillFromRef(
  xml: string,
  ref: XmlElement,
  deck: Deck,
  partName: string,
  rels: Map<string, Relation>,
  mode: 'fill' | 'bg'
): PaintFill {
  const element = styleAt(deck.theme, readNumber(attribute(ref.attrs, 'idx')), mode)
  if (element === undefined || deck.theme.xml === '') return NO_FILL
  const color = colorElement(xml, ref)
  const ctx: ColorContext = {
    theme: deck.theme,
    map: deck.map,
    placeholder: color === undefined ? deck.ctx.placeholder : rgbOf(xml, color, deck.ctx)
  }
  return paintFillElement(deck.theme.xml, element, ctx, deck, partName, rels, NO_FILL)
}

function styleAt(theme: Theme, idx: number, mode: 'fill' | 'bg'): XmlElement | undefined {
  if (idx <= 0) return undefined
  if (mode === 'bg' || idx >= 1001) {
    const index = idx >= 1001 ? idx - 1001 : idx - 1
    return theme.bgStyles[index]
  }
  return theme.fillStyles[idx - 1]
}

function gradientFill(xml: string, element: XmlElement, ctx: ColorContext): PaintFill {
  const list = direct(xml, element, 'gsLst')
  const stops = (list === undefined ? [] : childElements(xml, list.openEnd, list.closeStart))
    .filter((item) => item.name === 'gs')
    .map((stop) => {
      const color = colorElement(xml, stop)
      return {
        offset: readNumber(attribute(stop.attrs, 'pos')) / 100000,
        color: color === undefined ? '#000000' : cssColor(rgbOf(xml, color, ctx))
      }
    })
  const lin = direct(xml, element, 'lin')
  if (lin === undefined) return { kind: 'radial', stops }
  return {
    kind: 'linear',
    angle: readNumber(attribute(lin.attrs, 'ang')) / 60000,
    scaled: attribute(lin.attrs, 'scaled') !== '0',
    stops
  }
}

function pictureFill(
  xml: string,
  element: XmlElement,
  deck: Deck,
  partName: string,
  rels: Map<string, Relation>
): PaintFill {
  const blip = direct(xml, element, 'blip')
  const id = blip === undefined ? undefined : attribute(blip.attrs, 'embed', true)
  const target = id === undefined ? undefined : rels.get(id)?.target
  if (target === undefined) return NO_FILL
  const imagePath = resolvePart(partName, target)
  if (!deck.images.includes(imagePath)) deck.images.push(imagePath)
  const src = direct(xml, element, 'srcRect')
  const crop =
    src === undefined
      ? undefined
      : {
          left: readNumber(attribute(src.attrs, 'l')) / 100000,
          top: readNumber(attribute(src.attrs, 't')) / 100000,
          right: readNumber(attribute(src.attrs, 'r')) / 100000,
          bottom: readNumber(attribute(src.attrs, 'b')) / 100000
        }
  return { kind: 'picture', imagePath, crop }
}

function shapeLine(
  xml: string,
  spPr: XmlElement | undefined,
  style: XmlElement | undefined,
  deck: Deck
): PaintLine | undefined {
  if (spPr !== undefined) {
    const line = direct(xml, spPr, 'ln')
    if (line !== undefined) return lineFrom(xml, line, deck.ctx)
  }
  if (style === undefined) return undefined
  const ref = direct(xml, style, 'lnRef')
  if (ref === undefined) return undefined
  const themeLine = deck.theme.lineStyles[readNumber(attribute(ref.attrs, 'idx')) - 1]
  if (themeLine === undefined || deck.theme.xml === '') return undefined
  const color = colorElement(xml, ref)
  const ctx: ColorContext = {
    theme: deck.theme,
    map: deck.map,
    placeholder: color === undefined ? deck.ctx.placeholder : rgbOf(xml, color, deck.ctx)
  }
  return lineFrom(deck.theme.xml, themeLine, ctx)
}

function lineFrom(xml: string, line: XmlElement, ctx: ColorContext): PaintLine | undefined {
  if (direct(xml, line, 'noFill') !== undefined) return undefined
  const solid = direct(xml, line, 'solidFill')
  const colorEl = solid === undefined ? undefined : colorElement(xml, solid)
  const dash = direct(xml, line, 'prstDash')
  return {
    color:
      colorEl === undefined
        ? cssColor(schemeColor('tx1', ctx))
        : cssColor(rgbOf(xml, colorEl, ctx)),
    width: readNumber(attribute(line.attrs, 'w'), 9525),
    dash: dashOf(dash === undefined ? undefined : attribute(dash.attrs, 'val'))
  }
}

function dashOf(value: string | undefined): 'solid' | 'dash' | 'dot' {
  if (value === 'dot' || value === 'sysDot') return 'dot'
  if (value === undefined || value === 'solid') return 'solid'
  return 'dash'
}

function shapeShadow(
  xml: string,
  spPr: XmlElement | undefined,
  style: XmlElement | undefined,
  deck: Deck
): PaintShadow | undefined {
  if (spPr !== undefined) {
    const own = direct(xml, spPr, 'effectLst')
    if (own !== undefined) return shadowFrom(xml, own, deck.ctx)
  }
  if (style === undefined) return undefined
  const ref = direct(xml, style, 'effectRef')
  if (ref === undefined) return undefined
  const effect = deck.theme.effectStyles[readNumber(attribute(ref.attrs, 'idx')) - 1]
  if (effect === undefined || deck.theme.xml === '') return undefined
  const list = direct(deck.theme.xml, effect, 'effectLst')
  if (list === undefined) return undefined
  const color = colorElement(xml, ref)
  const ctx: ColorContext = {
    theme: deck.theme,
    map: deck.map,
    placeholder: color === undefined ? deck.ctx.placeholder : rgbOf(xml, color, deck.ctx)
  }
  return shadowFrom(deck.theme.xml, list, ctx)
}

function shadowFrom(xml: string, list: XmlElement, ctx: ColorContext): PaintShadow | undefined {
  const shadow = direct(xml, list, 'outerShdw')
  if (shadow === undefined) return undefined
  const colorEl = colorElement(xml, shadow)
  const distance = readNumber(attribute(shadow.attrs, 'dist'))
  const direction = (readNumber(attribute(shadow.attrs, 'dir')) / 60000) * (Math.PI / 180)
  return {
    color: colorEl === undefined ? 'rgba(0, 0, 0, 0.4)' : cssColor(rgbOf(xml, colorEl, ctx)),
    blur: readNumber(attribute(shadow.attrs, 'blurRad')),
    offsetX: Math.cos(direction) * distance,
    offsetY: Math.sin(direction) * distance
  }
}

function markersOf(xml: string): Marker[] {
  const tree = findDescendant(xml, 'spTree')
  if (tree === undefined) return []
  const found: Marker[] = []
  collectMarkers(xml, tree.openEnd, tree.closeStart, found)
  return found
}

function collectMarkers(xml: string, from: number, to: number, found: Marker[]): void {
  for (const element of childElements(xml, from, to)) {
    if (element.name === 'grpSp') {
      collectMarkers(xml, element.openEnd, element.closeStart, found)
      continue
    }
    const marker = markerOf(xml, element)
    if (marker !== undefined) found.push(marker)
  }
}

function markerOf(xml: string, element: XmlElement): Marker | undefined {
  const holder = directNames(xml, element, ['nvSpPr', 'nvPicPr', 'nvCxnSpPr', 'nvGraphicFramePr'])
  if (holder === undefined) return undefined
  const props = direct(xml, holder, 'nvPr')
  if (props === undefined) return undefined
  const placeholder = direct(xml, props, 'ph')
  if (placeholder === undefined) return undefined
  return {
    type: attribute(placeholder.attrs, 'type') ?? 'obj',
    idx: attribute(placeholder.attrs, 'idx'),
    element
  }
}

function findPrototype(markers: readonly Marker[], wanted: Marker): Marker | undefined {
  if (wanted.idx !== undefined) {
    const byIndex = markers.find((item) => item.idx === wanted.idx)
    if (byIndex !== undefined) return byIndex
  }
  const kind = alias(wanted.type)
  return markers.find((item) => alias(item.type) === kind)
}

function kindsOf(markers: readonly Marker[]): Set<string> {
  return new Set(markers.map((marker) => alias(marker.type)))
}

function alias(type: string): string {
  if (type === 'ctrTitle') return 'title'
  if (type === 'subTitle' || type === 'obj') return 'body'
  return type
}

function identityOf(xml: string, element: XmlElement): { id: string; name: string } | undefined {
  const holder = directNames(xml, element, [
    'nvSpPr',
    'nvPicPr',
    'nvCxnSpPr',
    'nvGrpSpPr',
    'nvGraphicFramePr'
  ])
  const node = findDescendant(
    xml,
    'cNvPr',
    (holder ?? element).openEnd,
    (holder ?? element).closeStart
  )
  if (node === undefined) return undefined
  const id = attribute(node.attrs, 'id')
  if (id === undefined) return undefined
  return { id, name: attribute(node.attrs, 'name') ?? '' }
}

function isHidden(xml: string, element: XmlElement): boolean {
  const holder = directNames(xml, element, [
    'nvSpPr',
    'nvPicPr',
    'nvCxnSpPr',
    'nvGrpSpPr',
    'nvGraphicFramePr'
  ])
  const node = findDescendant(
    xml,
    'cNvPr',
    (holder ?? element).openEnd,
    (holder ?? element).closeStart
  )
  return node !== undefined && attribute(node.attrs, 'hidden') === '1'
}

function direct(xml: string, element: XmlElement, name: string): XmlElement | undefined {
  return childElements(xml, element.openEnd, element.closeStart).find((item) => item.name === name)
}

function directNames(
  xml: string,
  element: XmlElement,
  names: readonly string[]
): XmlElement | undefined {
  return childElements(xml, element.openEnd, element.closeStart).find((item) =>
    names.includes(item.name)
  )
}
