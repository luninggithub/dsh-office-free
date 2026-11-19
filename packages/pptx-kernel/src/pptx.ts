import {
  EMPTY_XFRM,
  groupTransform,
  IDENTITY,
  mapBox,
  shapeXfrm,
  type Transform,
  type Xfrm
} from './geom.ts'
import { readRelations, resolvePart, type Relation } from './rels.ts'
import { buildSlideScene, type SlideScene } from './scene.ts'
import {
  attribute,
  childElements,
  decodeXml,
  escapeXml,
  findDescendant,
  parseAttributes,
  type XmlElement
} from './xml.ts'
import { readZip, writeZip } from './zip.ts'

/**
 * PPTX kernel.
 *
 * The package bytes are the source of truth. An edit rewrites only the text nodes of one shape;
 * every other part is copied through. Saving without an edit returns the original archive.
 * Grouped shapes honor `chOff` / `chExt`. `scene()` is the visual page: theme, master, layout,
 * fills, and run formatting. Text edits still go through `setShapeText`.
 */

const textDecoder = new TextDecoder()
const textEncoder = new TextEncoder()

export interface SlideSize {
  readonly width: number
  readonly height: number
}

interface ShapeBase {
  readonly id: string
  readonly name: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly text: string
}

export type SlideShape =
  | (ShapeBase & { readonly kind: 'text' | 'other' })
  | (ShapeBase & { readonly kind: 'picture'; readonly imagePath: string })

export interface SlidePage {
  readonly partName: string
  readonly shapes: readonly SlideShape[]
}

interface SlideEntry {
  readonly relationId: string
  readonly tag: string
  readonly partName: string
}

interface LocatedShape {
  readonly element: XmlElement
  readonly parent: Transform
  readonly local: Xfrm
}

/** EMU (English Metric Units) per CSS pixel at 96dpi. */
export const EMU_PER_PIXEL = 914400 / 96

export class PptxDocument {
  private readonly original: Uint8Array
  private readonly parts: Map<string, Uint8Array>
  private order: string[]
  private readonly edited = new Map<string, string>()
  private readonly scenes = new Map<number, SlideScene>()
  private dirty = false

  private constructor(
    original: Uint8Array,
    parts: Map<string, Uint8Array>,
    order: readonly string[]
  ) {
    this.original = original
    this.parts = parts
    this.order = [...order]
  }

  static async open(input: Uint8Array): Promise<PptxDocument> {
    const bytes = input.slice()
    const parts = await readZip(bytes)
    if (!parts.has('ppt/presentation.xml')) throw new Error('pptx: presentation.xml is missing')
    return new PptxDocument(bytes, parts, [...parts.keys()])
  }

  slideSize(): SlideSize {
    const xml = this.text('ppt/presentation.xml')
    const element = findDescendant(xml, 'sldSz')
    return {
      width: numberAttr(element?.attrs, 'cx'),
      height: numberAttr(element?.attrs, 'cy')
    }
  }

  slides(): SlidePage[] {
    const presentation = this.text('ppt/presentation.xml')
    const relations = readRelations(this.text('ppt/_rels/presentation.xml.rels'))
    const list = findDescendant(presentation, 'sldIdLst')
    if (list === undefined) return []
    const pages: SlidePage[] = []
    for (const slideId of childElements(presentation, list.openEnd, list.closeStart)) {
      if (slideId.name !== 'sldId') continue
      const relationId = attribute(slideId.attrs, 'id', true)
      const target = relationId === undefined ? undefined : relations.get(relationId)
      if (target === undefined || !target.type.endsWith('/slide')) continue
      const partName = resolvePart('ppt/presentation.xml', target.target)
      pages.push({
        partName,
        shapes: shapesOf(this.text(partName), partName, this.slideRelations(partName))
      })
    }
    return pages
  }

  /**
   * Visual page for one slide: theme colors and fonts, master and layout chrome,
   * shape fills, and run-level text.
   */
  scene(slideIndex: number): SlideScene {
    const cached = this.scenes.get(slideIndex)
    if (cached !== undefined) return cached
    const built = buildSlideScene((name) => this.optionalText(name), slideIndex)
    this.scenes.set(slideIndex, built)
    return built
  }

  /** Bytes of a package part, such as `ppt/media/image1.png`. */
  part(name: string): Uint8Array | undefined {
    const edited = this.edited.get(name)
    if (edited !== undefined) return textEncoder.encode(edited)
    const bytes = this.parts.get(name)
    return bytes === undefined ? undefined : bytes.slice()
  }

  /**
   * Replace the visible text of one shape. Characters that did not change stay in their
   * original runs, so a later run such as a unit keeps its size, weight, and color.
   */
  setShapeText(slideIndex: number, shapeId: string, text: string): void {
    this.writeShapeText(slideIndex, shapeId, (current) => distributeRuns(current, text))
  }

  /** Write one string per existing run. Run properties in the package are left untouched. */
  setShapeRuns(
    slideIndex: number,
    shapeId: string,
    paragraphs: readonly (readonly string[])[]
  ): void {
    this.writeShapeText(slideIndex, shapeId, () => paragraphs.map((runs) => [...runs]))
  }

  /**
   * Move and resize a shape. `frame` is the axis-aligned box in slide EMU, after group scaling.
   */
  setShapeFrame(
    slideIndex: number,
    shapeId: string,
    frame: {
      readonly x: number
      readonly y: number
      readonly width: number
      readonly height: number
    }
  ): void {
    this.rewriteShape(slideIndex, shapeId, (xml, located) => {
      const sx = located.parent.sx === 0 ? 1 : located.parent.sx
      const sy = located.parent.sy === 0 ? 1 : located.parent.sy
      return writeShapeBox(xml, located.element, {
        x: Math.round((frame.x - located.parent.x) / sx),
        y: Math.round((frame.y - located.parent.y) / sy),
        width: Math.max(1, Math.round(frame.width / sx)),
        height: Math.max(1, Math.round(frame.height / sy))
      })
    })
  }

  /** Crop a picture. Each edge is the fraction of the source image to hide, from 0 to 1. */
  setPictureCrop(
    slideIndex: number,
    shapeId: string,
    crop: {
      readonly left: number
      readonly top: number
      readonly right: number
      readonly bottom: number
    }
  ): void {
    this.rewriteShape(slideIndex, shapeId, (xml, located) =>
      writePictureCrop(xml, located.element, crop)
    )
  }

  /** Remove one shape from a slide. Shared pictures stay in the package. */
  deleteShape(slideIndex: number, shapeId: string): void {
    this.rewriteShape(slideIndex, shapeId, (xml, located) => {
      return xml.slice(0, located.element.start) + xml.slice(located.element.end)
    })
  }

  /**
   * Copy a shape onto `toSlide`. `dx` and `dy` are slide-EMU offsets from its current visual box.
   * Returns the new shape id.
   */
  duplicateShape(
    fromSlide: number,
    shapeId: string,
    toSlide: number,
    dx: number,
    dy: number
  ): string {
    const source = this.slides()[fromSlide]
    const target = this.slides()[toSlide]
    if (source === undefined) throw new Error(`pptx: slide ${fromSlide} does not exist`)
    if (target === undefined) throw new Error(`pptx: slide ${toSlide} does not exist`)
    const sourceXml = this.text(source.partName)
    const located = locateShape(sourceXml, shapeId)
    if (located === undefined) throw new Error(`pptx: shape ${shapeId} does not exist`)
    const visual = mapBox(
      located.parent,
      located.local.x,
      located.local.y,
      located.local.width,
      located.local.height
    )
    const targetXml = this.text(target.partName)
    const newId = String(nextShapeId(targetXml))
    let copy = replaceOwnId(sourceXml.slice(located.element.start, located.element.end), newId)
    copy = writeFragmentBox(copy, {
      x: Math.round(visual.x + dx),
      y: Math.round(visual.y + dy),
      width: Math.max(1, Math.round(visual.width)),
      height: Math.max(1, Math.round(visual.height))
    })
    const embed = attributeValue(copy, 'embed', true)
    if (embed !== undefined) {
      const relation = this.slideRelations(source.partName).get(embed)
      if (relation !== undefined) {
        const id = this.ensureRelationship(target.partName, relation.type, relation.target)
        copy = copy.replace(/(embed=")[^"]*(")/, `$1${id}$2`)
      }
    }
    const tree = findDescendant(this.text(target.partName), 'spTree')
    if (tree === undefined) throw new Error('pptx: slide has no shape tree')
    const next = this.text(target.partName)
    this.edited.set(
      target.partName,
      next.slice(0, tree.closeStart) + copy + next.slice(tree.closeStart)
    )
    this.dirty = true
    this.scenes.delete(fromSlide)
    this.scenes.delete(toSlide)
    return newId
  }

  /**
   * Append a blank slide, using the layout of `layoutFromIndex` so the deck template stays
   * in effect. Returns the new slide index.
   */
  addSlide(layoutFromIndex: number): number {
    const existing = this.slides()
    const index = existing.length
    const source =
      existing[Math.min(Math.max(layoutFromIndex, 0), Math.max(existing.length - 1, 0))]
    const layout =
      source === undefined ? this.anyLayoutTarget() : this.layoutTarget(source.partName)
    this.insertSlide(index, blankSlide(), slideRelationships(layout))
    return index
  }

  /** Insert an identical copy immediately after `index`. Returns the copy's index. */
  duplicateSlide(index: number): number {
    const source = this.slideEntries()[index]
    if (source === undefined) throw new Error(`pptx: slide ${index} does not exist`)
    const rels = this.optionalText(slideRelsName(source.partName)) ?? slideRelationships(undefined)
    this.insertSlide(index + 1, this.text(source.partName), rels)
    return index + 1
  }

  /**
   * Move the slide at `fromIndex` so it occupies `toIndex`.
   * Both indexes are positions in the current slide list.
   */
  moveSlide(fromIndex: number, toIndex: number): void {
    const entries = this.slideEntries()
    if (entries[fromIndex] === undefined) throw new Error(`pptx: slide ${fromIndex} does not exist`)
    if (entries[toIndex] === undefined) throw new Error(`pptx: slide ${toIndex} does not exist`)
    if (fromIndex === toIndex) return
    const next = entries.slice()
    const [moved] = next.splice(fromIndex, 1)
    if (moved === undefined) return
    next.splice(toIndex, 0, moved)
    this.remember(
      'ppt/presentation.xml',
      replaceSlideIds(
        this.text('ppt/presentation.xml'),
        next.map((entry) => entry.tag)
      )
    )
    this.dirty = true
    this.scenes.clear()
  }

  /** Remove one slide. A deck must keep its last slide. */
  deleteSlide(index: number): void {
    const entries = this.slideEntries()
    const entry = entries[index]
    if (entry === undefined) throw new Error(`pptx: slide ${index} does not exist`)
    if (entries.length <= 1) throw new Error('pptx: the only slide cannot be deleted')
    this.remember(
      'ppt/presentation.xml',
      replaceSlideIds(
        this.text('ppt/presentation.xml'),
        entries.filter((item) => item !== entry).map((item) => item.tag)
      )
    )
    this.remember(
      'ppt/_rels/presentation.xml.rels',
      removeRelationship(this.text('ppt/_rels/presentation.xml.rels'), entry.relationId)
    )
    const types = this.optionalText('[Content_Types].xml')
    if (types !== undefined) {
      this.remember('[Content_Types].xml', removeContentType(types, entry.partName))
    }
    this.forget(entry.partName)
    this.forget(slideRelsName(entry.partName))
    this.dirty = true
    this.scenes.clear()
  }

  private insertSlide(index: number, xml: string, rels: string): void {
    const file = this.nextSlideFile()
    this.remember(file.part, xml)
    this.remember(file.rels, rels)
    this.remember(
      'ppt/presentation.xml',
      insertSlideId(this.text('ppt/presentation.xml'), file.relationId, index)
    )
    this.remember(
      'ppt/_rels/presentation.xml.rels',
      insertRelationship(this.text('ppt/_rels/presentation.xml.rels'), file.relationId, file.target)
    )
    if (this.parts.has('[Content_Types].xml') || this.edited.has('[Content_Types].xml')) {
      this.remember(
        '[Content_Types].xml',
        insertContentType(this.text('[Content_Types].xml'), file.part)
      )
    }
    this.dirty = true
    this.scenes.clear()
  }

  private writeShapeText(
    slideIndex: number,
    shapeId: string,
    texts: (current: readonly (readonly string[])[]) => readonly (readonly string[])[]
  ): void {
    const slide = this.slides()[slideIndex]
    if (slide === undefined) throw new Error(`pptx: slide ${slideIndex} does not exist`)
    const xml = this.text(slide.partName)
    const shape = findShape(xml, shapeId)
    if (shape === undefined) throw new Error(`pptx: shape ${shapeId} does not exist`)
    const raw = xml.slice(shape.start, shape.end)
    const current = textGrid(raw)
    if (current.length === 0) throw new Error('pptx: shape has no text')
    const next = replaceShapeRuns(raw, texts(current))
    this.edited.set(slide.partName, xml.slice(0, shape.start) + next + xml.slice(shape.end))
    this.dirty = true
    this.scenes.delete(slideIndex)
  }

  private remember(name: string, xml: string): void {
    this.edited.set(name, xml)
    if (!this.order.includes(name)) this.order.push(name)
  }

  private forget(name: string): void {
    this.edited.delete(name)
    this.parts.delete(name)
    this.order = this.order.filter((item) => item !== name)
  }

  private rewriteShape(
    slideIndex: number,
    shapeId: string,
    change: (xml: string, located: LocatedShape) => string
  ): void {
    const slide = this.slides()[slideIndex]
    if (slide === undefined) throw new Error(`pptx: slide ${slideIndex} does not exist`)
    const xml = this.text(slide.partName)
    const located = locateShape(xml, shapeId)
    if (located === undefined) throw new Error(`pptx: shape ${shapeId} does not exist`)
    this.edited.set(slide.partName, change(xml, located))
    this.dirty = true
    this.scenes.delete(slideIndex)
  }

  private ensureRelationship(partName: string, type: string, target: string): string {
    const relsName = slideRelsName(partName)
    const xml = this.optionalText(relsName) ?? slideRelationships(undefined)
    const relations = readRelations(xml)
    for (const [id, relation] of relations) {
      if (relation.type === type && relation.target === target) return id
    }
    const id = nextRelationId(relations)
    this.remember(relsName, insertRelationship(xml, id, target, type))
    return id
  }

  private slideEntries(): SlideEntry[] {
    const presentation = this.text('ppt/presentation.xml')
    const relations = readRelations(this.text('ppt/_rels/presentation.xml.rels'))
    const list = findDescendant(presentation, 'sldIdLst')
    if (list === undefined) return []
    const entries: SlideEntry[] = []
    for (const slideId of childElements(presentation, list.openEnd, list.closeStart)) {
      if (slideId.name !== 'sldId') continue
      const relationId = attribute(slideId.attrs, 'id', true)
      const target = relationId === undefined ? undefined : relations.get(relationId)
      if (relationId === undefined || target === undefined || !target.type.endsWith('/slide')) {
        continue
      }
      entries.push({
        relationId,
        tag: presentation.slice(slideId.start, slideId.end),
        partName: resolvePart('ppt/presentation.xml', target.target)
      })
    }
    return entries
  }

  private layoutTarget(partName: string): string | undefined {
    for (const relation of this.slideRelations(partName).values()) {
      if (relation.type.endsWith('/slideLayout')) return relation.target
    }
    return this.anyLayoutTarget()
  }

  private anyLayoutTarget(): string | undefined {
    for (const name of this.order) {
      const file = /^ppt\/slideLayouts\/(slideLayout\d+\.xml)$/.exec(name)
      const base = file?.[1]
      if (base !== undefined) return `../slideLayouts/${base}`
    }
    return undefined
  }

  private nextSlideFile(): { part: string; rels: string; target: string; relationId: string } {
    let max = 0
    for (const name of new Set([...this.order, ...this.edited.keys()])) {
      const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(name)
      if (match !== null) max = Math.max(max, Number(match[1]))
    }
    const number = max + 1
    return {
      part: `ppt/slides/slide${number}.xml`,
      rels: `ppt/slides/_rels/slide${number}.xml.rels`,
      target: `slides/slide${number}.xml`,
      relationId: nextRelationId(readRelations(this.text('ppt/_rels/presentation.xml.rels')))
    }
  }

  /** Original bytes when nothing changed; otherwise a new package with untouched parts copied. */
  async save(): Promise<Uint8Array> {
    if (!this.dirty) return this.original.slice()
    return writeZip(
      this.order.map((name) => {
        const edited = this.edited.get(name)
        const data = edited === undefined ? this.parts.get(name) : textEncoder.encode(edited)
        if (data === undefined) throw new Error(`pptx: part ${name} disappeared`)
        return { name, data }
      })
    )
  }

  private text(name: string): string {
    const edited = this.edited.get(name)
    if (edited !== undefined) return edited
    const bytes = this.parts.get(name)
    if (bytes === undefined) throw new Error(`pptx: part ${name} is missing`)
    return textDecoder.decode(bytes)
  }

  private optionalText(name: string): string | undefined {
    if (!this.edited.has(name) && !this.parts.has(name)) return undefined
    return this.text(name)
  }

  private slideRelations(partName: string): Map<string, Relation> {
    const slash = partName.lastIndexOf('/')
    const rels = `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`
    return this.parts.has(rels) || this.edited.has(rels)
      ? readRelations(this.text(rels))
      : new Map()
  }
}

export function openPptx(input: Uint8Array): Promise<PptxDocument> {
  return PptxDocument.open(input)
}

function shapesOf(xml: string, partName: string, relations: Map<string, Relation>): SlideShape[] {
  const tree = findDescendant(xml, 'spTree')
  if (tree === undefined) return []
  return collectShapes(xml, partName, tree.openEnd, tree.closeStart, relations, IDENTITY)
}

function collectShapes(
  xml: string,
  partName: string,
  from: number,
  to: number,
  relations: Map<string, Relation>,
  origin: Transform
): SlideShape[] {
  const shapes: SlideShape[] = []
  for (const element of childElements(xml, from, to)) {
    if (element.name === 'grpSp') {
      const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM
      shapes.push(
        ...collectShapes(
          xml,
          partName,
          element.openEnd,
          element.closeStart,
          relations,
          groupTransform(origin, xfrm)
        )
      )
      continue
    }
    if (element.name !== 'sp' && element.name !== 'pic') continue
    const id = readShapeId(xml, element)
    if (id === undefined) continue
    const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM
    const box = mapBox(origin, xfrm.x, xfrm.y, xfrm.width, xfrm.height)
    const text = textOf(xml.slice(element.start, element.end))
    const embed = attributeValue(xml.slice(element.start, element.end), 'embed', true)
    const target = embed === undefined ? undefined : relations.get(embed)?.target
    const imagePath = target === undefined ? undefined : resolvePart(partName, target)
    const placed = {
      id,
      name: shapeName(xml, element),
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      text
    }
    if (element.name === 'pic' && imagePath !== undefined) {
      shapes.push({ kind: 'picture', ...placed, imagePath })
      continue
    }
    shapes.push({ kind: text.length > 0 ? 'text' : 'other', ...placed })
  }
  return shapes
}

function findShape(xml: string, shapeId: string): XmlElement | undefined {
  const tree = findDescendant(xml, 'spTree')
  if (tree === undefined) return undefined
  return findShapeIn(xml, tree.openEnd, tree.closeStart, shapeId)
}

function findShapeIn(
  xml: string,
  from: number,
  to: number,
  shapeId: string
): XmlElement | undefined {
  for (const element of childElements(xml, from, to)) {
    if (element.name === 'grpSp') {
      const nested = findShapeIn(xml, element.openEnd, element.closeStart, shapeId)
      if (nested !== undefined) return nested
      continue
    }
    if (
      (element.name === 'sp' || element.name === 'pic') &&
      readShapeId(xml, element) === shapeId
    ) {
      return element
    }
  }
  return undefined
}

function readShapeId(xml: string, element: XmlElement): string | undefined {
  const properties = childElements(xml, element.openEnd, element.closeStart)
  const holder = properties.find((item) => item.name === 'nvSpPr' || item.name === 'nvPicPr')
  const range = holder ?? element
  const marker = findDescendant(xml, 'cNvPr', range.openEnd, range.closeStart)
  return marker === undefined ? undefined : attribute(marker.attrs, 'id')
}

function shapeName(xml: string, element: XmlElement): string {
  const properties = childElements(xml, element.openEnd, element.closeStart)
  const holder = properties.find((item) => item.name === 'nvSpPr' || item.name === 'nvPicPr')
  const range = holder ?? element
  const marker = findDescendant(xml, 'cNvPr', range.openEnd, range.closeStart)
  return marker === undefined ? '' : (attribute(marker.attrs, 'name') ?? '')
}

function textOf(raw: string): string {
  return [...raw.matchAll(/<([A-Za-z_][\w.-]*):t(?:\s[^>]*)?>([\s\S]*?)<\/\1:t>/g)]
    .map((match) => decodeXml(match[2] ?? ''))
    .join('')
}

function textGrid(raw: string): string[][] {
  const body = findDescendant(raw, 'txBody')
  if (body === undefined) return []
  const grid: string[][] = []
  for (const paragraph of childElements(raw, body.openEnd, body.closeStart)) {
    if (paragraph.name !== 'p') continue
    const nodes = textNodesIn(raw, paragraph)
    grid.push(
      nodes.length === 0
        ? ['']
        : nodes.map((node) => decodeXml(raw.slice(node.openEnd, node.closeStart)))
    )
  }
  return grid
}

function replaceShapeRuns(raw: string, paragraphs: readonly (readonly string[])[]): string {
  const body = findDescendant(raw, 'txBody')
  if (body === undefined) throw new Error('pptx: shape has no text')
  const blocks = childElements(raw, body.openEnd, body.closeStart).filter(
    (item) => item.name === 'p'
  )
  const edits: Array<{ start: number; end: number; xml: string }> = []
  blocks.forEach((paragraph, index) => {
    const wanted = paragraphs[index] ?? ['']
    const nodes = textNodesIn(raw, paragraph)
    if (nodes.length === 0) {
      const text = wanted.join('')
      if (text.length === 0) return
      edits.push({
        start: paragraph.openEnd,
        end: paragraph.openEnd,
        xml: runXml(paragraph.prefix || 'a', text)
      })
      return
    }
    nodes.forEach((node, runIndex) => {
      const extra = runIndex === nodes.length - 1 ? wanted.slice(nodes.length).join('') : ''
      edits.push({
        start: node.start,
        end: node.end,
        xml: textElementXml(node.prefix, node.attrs, `${wanted[runIndex] ?? ''}${extra}`)
      })
    })
  })
  const overflow = paragraphs.slice(blocks.length)
  const last = blocks[blocks.length - 1]
  if (overflow.length > 0 && last !== undefined) {
    const text = overflow.map((runs) => runs.join('')).join('\n')
    if (text.length > 0) {
      edits.push({
        start: last.closeStart,
        end: last.closeStart,
        xml: runXml(last.prefix || 'a', text)
      })
    }
  }
  let next = raw
  for (const edit of edits.toSorted((a, b) => b.start - a.start)) {
    next = next.slice(0, edit.start) + edit.xml + next.slice(edit.end)
  }
  return next
}

function textNodesIn(xml: string, element: XmlElement): XmlElement[] {
  const found: XmlElement[] = []
  const walk = (from: number, to: number): void => {
    for (const child of childElements(xml, from, to)) {
      if (child.name === 't') found.push(child)
      else walk(child.openEnd, child.closeStart)
    }
  }
  walk(element.openEnd, element.closeStart)
  return found
}

function textElementXml(prefix: string, attrs: string, text: string): string {
  const tag = prefix.length > 0 ? `${prefix}:t` : 't'
  return `<${tag}${withPreserve(attrs, text)}>${escapeXml(text)}</${tag}>`
}

function runXml(prefix: string, text: string): string {
  const run = prefix.length > 0 ? prefix : 'a'
  return `<${run}:r>${textElementXml(run, '', text)}</${run}:r>`
}

/**
 * Keep unchanged characters in the runs that already owned them. Inserted characters take
 * the run before the caret when the edit sits on a run boundary.
 */
function distributeRuns(previous: readonly (readonly string[])[], nextPlain: string): string[][] {
  const nextLines = nextPlain.split('\n')
  if (previous.length === 0) return nextLines.map((line) => [line])
  if (nextLines.length === previous.length) {
    return previous.map((runs, index) => distributeParagraph(runs, nextLines[index] ?? ''))
  }
  const head = previous.slice(0, -1)
  const tail = previous[previous.length - 1] ?? ['']
  return [
    ...head.map((runs, index) => distributeParagraph(runs, nextLines[index] ?? '')),
    distributeParagraph(tail, nextLines.slice(head.length).join('\n'))
  ]
}

function distributeParagraph(runs: readonly string[], next: string): string[] {
  if (runs.length <= 1) return [next]
  const previous = runs.join('')
  let prefix = 0
  const shared = Math.min(previous.length, next.length)
  while (prefix < shared && previous[prefix] === next[prefix]) prefix += 1
  let suffix = 0
  while (
    suffix < previous.length - prefix &&
    suffix < next.length - prefix &&
    previous[previous.length - 1 - suffix] === next[next.length - 1 - suffix]
  ) {
    suffix += 1
  }
  const owners = runOwners(runs)
  const insertAt = prefix >= previous.length ? runs.length - 1 : boundaryRun(owners, prefix)
  const buckets = runs.map(() => '')
  const push = (run: number, char: string): void => {
    buckets[run] = `${buckets[run] ?? ''}${char}`
  }
  for (let index = 0; index < prefix; index += 1) push(owners[index] ?? 0, next[index] ?? '')
  for (let index = prefix; index < next.length - suffix; index += 1)
    push(insertAt, next[index] ?? '')
  for (let index = 0; index < suffix; index += 1) {
    const oldIndex = previous.length - suffix + index
    push(owners[oldIndex] ?? runs.length - 1, next[next.length - suffix + index] ?? '')
  }
  return buckets
}

function runOwners(runs: readonly string[]): number[] {
  const owners: number[] = []
  runs.forEach((text, index) => {
    for (let cursor = 0; cursor < text.length; cursor += 1) owners.push(index)
  })
  return owners
}

function boundaryRun(owners: readonly number[], prefix: number): number {
  const at = owners[prefix] ?? 0
  if (prefix === 0) return at
  const before = owners[prefix - 1] ?? at
  return before === at ? at : before
}

function blankSlide(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`
}

function locateShape(xml: string, shapeId: string): LocatedShape | undefined {
  const tree = findDescendant(xml, 'spTree')
  if (tree === undefined) return undefined
  return walkShapes(xml, tree.openEnd, tree.closeStart, shapeId, IDENTITY)
}

function walkShapes(
  xml: string,
  from: number,
  to: number,
  shapeId: string,
  parent: Transform
): LocatedShape | undefined {
  for (const element of childElements(xml, from, to)) {
    if (element.name === 'grpSp') {
      const found = walkShapes(
        xml,
        element.openEnd,
        element.closeStart,
        shapeId,
        groupTransform(parent, shapeXfrm(xml, element) ?? EMPTY_XFRM)
      )
      if (found !== undefined) return found
      continue
    }
    if (
      (element.name === 'sp' || element.name === 'pic') &&
      readShapeId(xml, element) === shapeId
    ) {
      return { element, parent, local: shapeXfrm(xml, element) ?? EMPTY_XFRM }
    }
  }
  return undefined
}

function xfrmElement(xml: string, element: XmlElement): XmlElement | undefined {
  const children = childElements(xml, element.openEnd, element.closeStart)
  const props = children.find((item) => item.name === 'spPr' || item.name === 'grpSpPr')
  if (props !== undefined) return findDescendant(xml, 'xfrm', props.openEnd, props.closeStart)
  return children.find((item) => item.name === 'xfrm')
}

function writeShapeBox(
  xml: string,
  element: XmlElement,
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
): string {
  const xfrm = xfrmElement(xml, element)
  const off = `<a:off x="${box.x}" y="${box.y}"/>`
  const ext = `<a:ext cx="${box.width}" cy="${box.height}"/>`
  if (xfrm === undefined) {
    const props = childElements(xml, element.openEnd, element.closeStart).find(
      (item) => item.name === 'spPr'
    )
    const tag = `<a:xfrm>${off}${ext}</a:xfrm>`
    const at = props?.openEnd ?? element.closeStart
    return xml.slice(0, at) + tag + xml.slice(at)
  }
  if (xfrm.openEnd === xfrm.end) {
    return xml.slice(0, xfrm.start) + `<a:xfrm>${off}${ext}</a:xfrm>` + xml.slice(xfrm.end)
  }
  const children = childElements(xml, xfrm.openEnd, xfrm.closeStart)
  const parts: string[] = []
  let wroteOff = false
  let wroteExt = false
  for (const child of children) {
    if (child.name === 'off') {
      parts.push(off)
      wroteOff = true
      continue
    }
    if (child.name === 'ext') {
      parts.push(ext)
      wroteExt = true
      continue
    }
    parts.push(xml.slice(child.start, child.end))
  }
  if (!wroteOff) parts.unshift(off)
  if (!wroteExt) parts.splice(1, 0, ext)
  return xml.slice(0, xfrm.openEnd) + parts.join('') + xml.slice(xfrm.closeStart)
}

function writeFragmentBox(
  fragment: string,
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
): string {
  const root = childElements(fragment)[0]
  if (root === undefined) return fragment
  return writeShapeBox(fragment, root, box)
}

function writePictureCrop(
  xml: string,
  element: XmlElement,
  crop: {
    readonly left: number
    readonly top: number
    readonly right: number
    readonly bottom: number
  }
): string {
  const fill = childElements(xml, element.openEnd, element.closeStart).find(
    (item) => item.name === 'blipFill'
  )
  if (fill === undefined) throw new Error('pptx: shape has no picture')
  const tag = `<a:srcRect l="${cropEdge(crop.left)}" t="${cropEdge(crop.top)}" r="${cropEdge(crop.right)}" b="${cropEdge(crop.bottom)}"/>`
  const existing = childElements(xml, fill.openEnd, fill.closeStart).find(
    (item) => item.name === 'srcRect'
  )
  if (existing !== undefined) return xml.slice(0, existing.start) + tag + xml.slice(existing.end)
  return xml.slice(0, fill.closeStart) + tag + xml.slice(fill.closeStart)
}

function cropEdge(value: number): number {
  return Math.round(Math.min(0.9, Math.max(0, value)) * 100000)
}

function replaceOwnId(fragment: string, id: string): string {
  return fragment.replace(/(<[^>]*\bcNvPr\b[^>]*\bid=")[^"]*(")/, `$1${id}$2`)
}

function nextShapeId(xml: string): number {
  let max = 1
  for (const match of xml.matchAll(/<[^>]*\bcNvPr\b[^>]*\bid="(\d+)"/g)) {
    max = Math.max(max, Number(match[1] ?? 0))
  }
  return max + 1
}

function slideRelsName(partName: string): string {
  const slash = partName.lastIndexOf('/')
  return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`
}

function slideRelationships(layout: string | undefined): string {
  const relationship =
    layout === undefined
      ? ''
      : `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${escapeXml(layout)}"/>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`
}

function replaceSlideIds(presentation: string, tags: readonly string[]): string {
  const list = findDescendant(presentation, 'sldIdLst')
  if (list === undefined) return presentation
  return presentation.slice(0, list.openEnd) + tags.join('') + presentation.slice(list.closeStart)
}

function removeRelationship(xml: string, id: string): string {
  const root = findDescendant(xml, 'Relationships')
  if (root === undefined) return xml
  const matches = childElements(xml, root.openEnd, root.closeStart).filter(
    (item) => item.name === 'Relationship' && attribute(item.attrs, 'Id') === id
  )
  let next = xml
  for (const item of matches.toReversed()) next = next.slice(0, item.start) + next.slice(item.end)
  return next
}

function removeContentType(xml: string, partName: string): string {
  const types = findDescendant(xml, 'Types')
  if (types === undefined) return xml
  const matches = childElements(xml, types.openEnd, types.closeStart).filter(
    (item) => item.name === 'Override' && attribute(item.attrs, 'PartName') === `/${partName}`
  )
  let next = xml
  for (const item of matches.toReversed()) next = next.slice(0, item.start) + next.slice(item.end)
  return next
}

function insertSlideId(presentation: string, relationId: string, index: number): string {
  const list = findDescendant(presentation, 'sldIdLst')
  const id = nextSlideId(presentation)
  const prefix = relationPrefix(presentation)
  const tag = `<p:sldId id="${id}" ${prefix}:id="${relationId}"/>`
  if (list === undefined) {
    const root = findDescendant(presentation, 'presentation')
    const at = root?.openEnd ?? presentation.length
    return `${presentation.slice(0, at)}<p:sldIdLst>${tag}</p:sldIdLst>${presentation.slice(at)}`
  }
  const ids = childElements(presentation, list.openEnd, list.closeStart).filter(
    (item) => item.name === 'sldId'
  )
  const anchor = index <= 0 ? undefined : ids[index - 1]
  const at = anchor?.end ?? (index <= 0 ? list.openEnd : list.closeStart)
  return presentation.slice(0, at) + tag + presentation.slice(at)
}

function nextSlideId(presentation: string): number {
  let max = 255
  const list = findDescendant(presentation, 'sldIdLst')
  if (list === undefined) return 256
  for (const item of childElements(presentation, list.openEnd, list.closeStart)) {
    if (item.name !== 'sldId') continue
    const id = Number(attribute(item.attrs, 'id') ?? 0)
    if (Number.isFinite(id)) max = Math.max(max, id)
  }
  return max + 1
}

function relationPrefix(presentation: string): string {
  const list = findDescendant(presentation, 'sldIdLst')
  if (list === undefined) return 'r'
  for (const item of childElements(presentation, list.openEnd, list.closeStart)) {
    if (item.name !== 'sldId') continue
    const prefixed = parseAttributes(item.attrs).find(
      (attr) => attr.name === 'id' && attr.prefix !== ''
    )
    if (prefixed !== undefined) return prefixed.prefix
  }
  return 'r'
}

function insertRelationship(
  xml: string,
  id: string,
  target: string,
  type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
): string {
  const tag = `<Relationship Id="${id}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"/>`
  const close = xml.lastIndexOf('</Relationships>')
  if (close < 0) throw new Error('pptx: presentation relationships are missing')
  return xml.slice(0, close) + tag + xml.slice(close)
}

function insertContentType(xml: string, partName: string): string {
  const part = `/${partName}`
  if (xml.includes(`PartName="${part}"`)) return xml
  const tag = `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  const close = xml.lastIndexOf('</Types>')
  if (close < 0) return xml
  return xml.slice(0, close) + tag + xml.slice(close)
}

function nextRelationId(relations: Map<string, { type: string; target: string }>): string {
  let max = 0
  for (const id of relations.keys()) {
    const number = Number(/^rId(\d+)$/.exec(id)?.[1] ?? 0)
    if (number > max) max = number
  }
  return `rId${max + 1}`
}

function withPreserve(attrs: string, text: string): string {
  if (!/^\s|\s$|\n/.test(text) || attrs.includes('xml:space')) return attrs
  return `${attrs} xml:space="preserve"`
}

function attributeValue(raw: string, name: string, prefixed: boolean): string | undefined {
  const pattern = prefixed
    ? new RegExp(`(?:\\s)([A-Za-z_][\\w.-]*):${name}="([^"]*)"`)
    : new RegExp(`(?:\\s)${name}="([^"]*)"`)
  const match = pattern.exec(raw)
  if (match === null) return undefined
  return decodeXml(prefixed ? (match[2] ?? '') : (match[1] ?? ''))
}

function numberAttr(attrs: string | undefined, name: string): number {
  if (attrs === undefined) return 0
  const value = attribute(attrs, name)
  const parsed = value === undefined ? 0 : Number(value)
  return Number.isFinite(parsed) ? parsed : 0
}
