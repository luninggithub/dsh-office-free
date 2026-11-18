import { EMPTY_XFRM, groupTransform, IDENTITY, mapBox, shapeXfrm, type Transform } from './geom.ts'
import { readRelations, resolvePart, type Relation } from './rels.ts'
import { buildSlideScene, type SlideScene } from './scene.ts'
import {
  attribute,
  childElements,
  decodeXml,
  escapeXml,
  findDescendant,
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

/** EMU (English Metric Units) per CSS pixel at 96dpi. */
export const EMU_PER_PIXEL = 914400 / 96

export class PptxDocument {
  private readonly original: Uint8Array
  private readonly parts: Map<string, Uint8Array>
  private readonly order: readonly string[]
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
    this.order = order
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
   * Replace the visible text of one shape. The first text node keeps its run properties and
   * receives the whole string; later text nodes in that shape become empty.
   */
  setShapeText(slideIndex: number, shapeId: string, text: string): void {
    const slide = this.slides()[slideIndex]
    if (slide === undefined) throw new Error(`pptx: slide ${slideIndex} does not exist`)
    const xml = this.text(slide.partName)
    const shape = findShape(xml, shapeId)
    if (shape === undefined) throw new Error(`pptx: shape ${shapeId} does not exist`)
    const next = replaceShapeText(xml.slice(shape.start, shape.end), text)
    this.edited.set(slide.partName, xml.slice(0, shape.start) + next + xml.slice(shape.end))
    this.dirty = true
    this.scenes.delete(slideIndex)
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

function replaceShapeText(raw: string, text: string): string {
  const pattern = /<([A-Za-z_][\w.-]*):t(\s[^>]*)?>([\s\S]*?)<\/\1:t>/g
  const matches = [...raw.matchAll(pattern)]
  if (matches.length === 0) return insertTextRun(raw, text)
  let seen = false
  return raw.replace(pattern, (_full, prefix: string, attrs: string | undefined) => {
    const attributes = attrs ?? ''
    if (seen) return `<${prefix}:t${attributes}></${prefix}:t>`
    seen = true
    return `<${prefix}:t${withPreserve(attributes, text)}>${escapeXml(text)}</${prefix}:t>`
  })
}

function insertTextRun(raw: string, text: string): string {
  const body = /<([A-Za-z_][\w.-]*):txBody\b/.exec(raw)
  const paragraph = /<([A-Za-z_][\w.-]*):p(\s[^>]*)?>/.exec(raw)
  if (body === null || paragraph === null) throw new Error('pptx: shape has no text')
  const prefix = body[1] ?? 'a'
  const node = `<${prefix}:r><${prefix}:t${withPreserve('', text)}>${escapeXml(text)}</${prefix}:t></${prefix}:r>`
  const at = paragraph.index + paragraph[0].length
  return raw.slice(0, at) + node + raw.slice(at)
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
