import { readNumber } from './geom.ts'
import { attribute, childElements, findDescendant, type XmlElement } from './xml.ts'

/**
 * Theme colors and fonts.
 * Modifiers are applied in sRGB HSL. Office uses a linear-light variant, so gradient stops
 * are close to PowerPoint rather than byte-identical.
 */

export interface RGB {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

export interface FontSet {
  readonly latin: string
  readonly ea: string
  readonly cs: string
  readonly hans: string
}

export interface Theme {
  readonly xml: string
  readonly colors: Readonly<Record<string, RGB>>
  readonly major: FontSet
  readonly minor: FontSet
  readonly fillStyles: readonly XmlElement[]
  readonly lineStyles: readonly XmlElement[]
  readonly effectStyles: readonly XmlElement[]
  readonly bgStyles: readonly XmlElement[]
}

export type ColorMap = Readonly<Record<string, string>>

export interface ColorContext {
  readonly theme: Theme
  readonly map: ColorMap
  readonly placeholder: RGB | undefined
}

const MAP_KEYS = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink'
] as const

const DEFAULT_HEX: Readonly<Record<string, string>> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '1F497D',
  lt2: 'EEECE1',
  accent1: '4F81BD',
  accent2: 'C0504D',
  accent3: '9BBB59',
  accent4: '8064A2',
  accent5: '4BACC6',
  accent6: 'F79646',
  hlink: '0000FF',
  folHlink: '800080'
}

const DEFAULT_FONT: FontSet = { latin: 'Calibri', ea: '', cs: '', hans: '宋体' }

const PRESET: Readonly<Record<string, string>> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080'
}

const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr'])

export function defaultColorMap(): ColorMap {
  return {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'lt2',
    tx2: 'dk2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink'
  }
}

export function colorMapFrom(attrs: string | undefined): ColorMap | undefined {
  if (attrs === undefined) return undefined
  const map: Record<string, string> = { ...defaultColorMap() }
  let seen = false
  for (const key of MAP_KEYS) {
    const value = attribute(attrs, key)
    if (value === undefined) continue
    map[key] = value
    seen = true
  }
  return seen ? map : undefined
}

export function readTheme(xml: string | undefined): Theme {
  const colors = defaultPalette()
  if (xml === undefined) return emptyTheme(colors)
  const scheme = findDescendant(xml, 'clrScheme')
  if (scheme !== undefined) {
    for (const entry of childElements(xml, scheme.openEnd, scheme.closeStart)) {
      const paint = readSchemeEntry(xml, entry)
      if (paint !== undefined) colors[entry.name] = paint
    }
  }
  const fonts = findDescendant(xml, 'fontScheme')
  const format = findDescendant(xml, 'fmtScheme')
  return {
    xml,
    colors,
    major: readFonts(xml, fonts ? direct(xml, fonts, 'majorFont') : undefined),
    minor: readFonts(xml, fonts ? direct(xml, fonts, 'minorFont') : undefined),
    fillStyles: styleList(xml, format, 'fillStyleLst'),
    lineStyles: styleList(xml, format, 'lnStyleLst'),
    effectStyles: styleList(xml, format, 'effectStyleLst'),
    bgStyles: styleList(xml, format, 'bgFillStyleLst')
  }
}

export function colorElement(xml: string, parent: XmlElement): XmlElement | undefined {
  return childElements(xml, parent.openEnd, parent.closeStart).find((item) =>
    COLOR_TAGS.has(item.name)
  )
}

export function schemeColor(value: string, ctx: ColorContext): RGB {
  if (value === 'phClr') return ctx.placeholder ?? { r: 0, g: 0, b: 0, a: 1 }
  const mapped = ctx.map[value] ?? value
  return ctx.theme.colors[mapped] ?? ctx.theme.colors[value] ?? { r: 0, g: 0, b: 0, a: 1 }
}

export function rgbOf(xml: string, element: XmlElement, ctx: ColorContext): RGB {
  const value = attribute(element.attrs, 'val') ?? ''
  let base: RGB
  if (element.name === 'srgbClr') base = hexRgb(value)
  else if (element.name === 'schemeClr') base = schemeColor(value, ctx)
  else if (element.name === 'sysClr') base = hexRgb(attribute(element.attrs, 'lastClr') ?? '000000')
  else if (element.name === 'prstClr') base = hexRgb(PRESET[value] ?? '000000')
  else if (element.name === 'scrgbClr') base = scrgb(element.attrs)
  else base = { r: 0, g: 0, b: 0, a: 1 }
  return applyMods(base, childElements(xml, element.openEnd, element.closeStart))
}

export function cssColor(color: RGB): string {
  const r = clampByte(color.r)
  const g = clampByte(color.g)
  const b = clampByte(color.b)
  const alpha = Math.round(clamp01(color.a) * 1000) / 1000
  if (alpha >= 0.999) return `#${hex(r)}${hex(g)}${hex(b)}`
  return `rgba(${r}, ${g}, ${b}, ${alpha})`
}

function direct(xml: string, element: XmlElement, name: string): XmlElement | undefined {
  return childElements(xml, element.openEnd, element.closeStart).find((item) => item.name === name)
}

function emptyTheme(colors: Record<string, RGB>): Theme {
  return {
    xml: '',
    colors,
    major: DEFAULT_FONT,
    minor: DEFAULT_FONT,
    fillStyles: [],
    lineStyles: [],
    effectStyles: [],
    bgStyles: []
  }
}

function defaultPalette(): Record<string, RGB> {
  const colors: Record<string, RGB> = {}
  for (const [name, value] of Object.entries(DEFAULT_HEX)) colors[name] = hexRgb(value)
  return colors
}

function readSchemeEntry(xml: string, element: XmlElement): RGB | undefined {
  const inner = colorElement(xml, element)
  if (inner === undefined) return undefined
  if (inner.name === 'sysClr') return hexRgb(attribute(inner.attrs, 'lastClr') ?? '000000')
  if (inner.name === 'srgbClr') return hexRgb(attribute(inner.attrs, 'val') ?? '000000')
  return undefined
}

function readFonts(xml: string, element: XmlElement | undefined): FontSet {
  if (element === undefined) return DEFAULT_FONT
  const hans = childElements(xml, element.openEnd, element.closeStart).find(
    (item) => item.name === 'font' && attribute(item.attrs, 'script') === 'Hans'
  )
  return {
    latin: face(xml, element, 'latin') || DEFAULT_FONT.latin,
    ea: face(xml, element, 'ea'),
    cs: face(xml, element, 'cs'),
    hans: (hans === undefined ? undefined : attribute(hans.attrs, 'typeface')) || DEFAULT_FONT.hans
  }
}

function face(xml: string, element: XmlElement, name: string): string {
  const node = direct(xml, element, name)
  if (node === undefined) return ''
  return attribute(node.attrs, 'typeface') ?? ''
}

function styleList(
  xml: string,
  format: XmlElement | undefined,
  name: string
): readonly XmlElement[] {
  if (format === undefined) return []
  const list = direct(xml, format, name)
  if (list === undefined) return []
  return childElements(xml, list.openEnd, list.closeStart)
}

function scrgb(attrs: string): RGB {
  const channel = (name: string) => (readNumber(attribute(attrs, name)) / 100000) * 255
  return { r: channel('r'), g: channel('g'), b: channel('b'), a: 1 }
}

function applyMods(input: RGB, mods: readonly XmlElement[]): RGB {
  const color = { r: input.r, g: input.g, b: input.b, a: input.a }
  for (const mod of mods) {
    if (COLOR_TAGS.has(mod.name)) continue
    const amount = readNumber(attribute(mod.attrs, 'val')) / 100000
    if (mod.name === 'alpha') color.a *= amount
    else if (mod.name === 'alphaOff') color.a = clamp01(color.a + amount)
    else if (mod.name === 'shade') {
      color.r *= amount
      color.g *= amount
      color.b *= amount
    } else if (mod.name === 'tint') {
      color.r = color.r * amount + 255 * (1 - amount)
      color.g = color.g * amount + 255 * (1 - amount)
      color.b = color.b * amount + 255 * (1 - amount)
    } else if (mod.name === 'lumMod' || mod.name === 'lumOff' || mod.name === 'satMod') {
      const next = hslAdjust(color, mod.name, amount)
      color.r = next.r
      color.g = next.g
      color.b = next.b
      color.a = next.a
    }
  }
  return color
}

function hslAdjust(color: RGB, name: string, amount: number): RGB {
  const hsl = rgbToHsl(color)
  let { s, l } = hsl
  if (name === 'lumMod') l *= amount
  if (name === 'lumOff') l += amount
  if (name === 'satMod') s *= amount
  const next = hslToRgb(hsl.h, clamp01(s), clamp01(l))
  return { ...next, a: color.a }
}

function rgbToHsl(color: RGB): { h: number; s: number; l: number } {
  const r = color.r / 255
  const g = color.g / 255
  const b = color.b / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return { h: 0, s: 0, l }
  const delta = max - min
  const s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min)
  let h = 0
  if (max === r) h = (g - b) / delta + (g < b ? 6 : 0)
  else if (max === g) h = (b - r) / delta + 2
  else h = (r - g) / delta + 4
  return { h: h * 60, s, l }
}

function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  if (s === 0) {
    const gray = l * 255
    return { r: gray, g: gray, b: gray }
  }
  const chroma = (1 - Math.abs(2 * l - 1)) * s
  const hp = h / 60
  const x = chroma * (1 - Math.abs((hp % 2) - 1))
  const m = l - chroma / 2
  let r = 0
  let g = 0
  let b = 0
  if (hp < 1) {
    r = chroma
    g = x
  } else if (hp < 2) {
    r = x
    g = chroma
  } else if (hp < 3) {
    g = chroma
    b = x
  } else if (hp < 4) {
    g = x
    b = chroma
  } else if (hp < 5) {
    r = x
    b = chroma
  } else {
    r = chroma
    b = x
  }
  return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255 }
}

function hexRgb(text: string): RGB {
  const value = text.trim()
  const r = Number.parseInt(value.slice(0, 2), 16)
  const g = Number.parseInt(value.slice(2, 4), 16)
  const b = Number.parseInt(value.slice(4, 6), 16)
  const a = value.length >= 8 ? Number.parseInt(value.slice(6, 8), 16) / 255 : 1
  return {
    r: Number.isFinite(r) ? r : 0,
    g: Number.isFinite(g) ? g : 0,
    b: Number.isFinite(b) ? b : 0,
    a: Number.isFinite(a) ? a : 1
  }
}

function clampByte(value: number): number {
  return Math.min(255, Math.max(0, Math.round(value)))
}

function clamp01(value: number): number {
  return Math.min(1, Math.max(0, value))
}

function hex(value: number): string {
  return clampByte(value).toString(16).padStart(2, '0')
}
