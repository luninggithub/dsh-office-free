/** Quote-aware XML element scan. Offsets refer to the original document string. */

export interface XmlElement {
  readonly prefix: string
  readonly name: string
  readonly attrs: string
  readonly start: number
  readonly openEnd: number
  readonly closeStart: number
  readonly end: number
}

interface Tag {
  readonly kind: 'open' | 'close' | 'empty'
  readonly prefix: string
  readonly name: string
  readonly attrs: string
  readonly start: number
  readonly end: number
}

/** Direct child elements of the range `[from, to)`. */
export function childElements(xml: string, from = 0, to = xml.length): XmlElement[] {
  const result: XmlElement[] = []
  let index = from
  while (index < to) {
    const start = xml.indexOf('<', index)
    if (start < 0 || start >= to) break
    const tag = readTag(xml, start)
    if (tag === undefined || tag.kind === 'close') {
      index = start + 1
      continue
    }
    if (tag.kind === 'empty') {
      result.push({
        prefix: tag.prefix,
        name: tag.name,
        attrs: tag.attrs,
        start: tag.start,
        openEnd: tag.end,
        closeStart: tag.end,
        end: tag.end
      })
      index = tag.end
      continue
    }
    let depth = 1
    let cursor = tag.end
    let closeStart = tag.end
    while (cursor < xml.length && depth > 0) {
      const next = xml.indexOf('<', cursor)
      if (next < 0) break
      const inner = readTag(xml, next)
      if (inner === undefined) {
        cursor = next + 1
        continue
      }
      if (inner.prefix === tag.prefix && inner.name === tag.name) {
        if (inner.kind === 'open') depth += 1
        else if (inner.kind === 'close') {
          depth -= 1
          if (depth === 0) closeStart = inner.start
        }
      }
      cursor = inner.end
    }
    result.push({
      prefix: tag.prefix,
      name: tag.name,
      attrs: tag.attrs,
      start: tag.start,
      openEnd: tag.end,
      closeStart,
      end: cursor
    })
    index = cursor
  }
  return result
}

export function findDescendant(
  xml: string,
  name: string,
  from = 0,
  to = xml.length
): XmlElement | undefined {
  for (const element of childElements(xml, from, to)) {
    if (element.name === name) return element
    const nested = findDescendant(xml, name, element.openEnd, element.closeStart)
    if (nested !== undefined) return nested
  }
  return undefined
}

export interface XmlAttribute {
  readonly prefix: string
  readonly name: string
  readonly value: string
}

export function parseAttributes(attrs: string): XmlAttribute[] {
  const parsed: XmlAttribute[] = []
  let index = 0
  while (index < attrs.length) {
    while (index < attrs.length && /\s/.test(attrs[index] ?? '')) index += 1
    if (index >= attrs.length) break
    const nameStart = index
    while (index < attrs.length && /[A-Za-z0-9_.:-]/.test(attrs[index] ?? '')) index += 1
    const full = attrs.slice(nameStart, index)
    while (index < attrs.length && /\s/.test(attrs[index] ?? '')) index += 1
    if (attrs[index] !== '=') continue
    index += 1
    while (index < attrs.length && /\s/.test(attrs[index] ?? '')) index += 1
    const quote = attrs[index]
    if (quote !== '"' && quote !== "'") continue
    const valueStart = index + 1
    const valueEnd = attrs.indexOf(quote, valueStart)
    if (valueEnd < 0) break
    const colon = full.indexOf(':')
    parsed.push({
      prefix: colon === -1 ? '' : full.slice(0, colon),
      name: colon === -1 ? full : full.slice(colon + 1),
      value: decodeXml(attrs.slice(valueStart, valueEnd))
    })
    index = valueEnd + 1
  }
  return parsed
}

/** Unprefixed attribute, or the prefixed one when `prefixed` is set. */
export function attribute(attrs: string, name: string, prefixed = false): string | undefined {
  return parseAttributes(attrs).find(
    (item) => item.name === name && (prefixed ? item.prefix !== '' : item.prefix === '')
  )?.value
}

export function decodeXml(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (entity, token: string) => {
    if (token === 'amp') return '&'
    if (token === 'lt') return '<'
    if (token === 'gt') return '>'
    if (token === 'quot') return '"'
    if (token === 'apos') return "'"
    const codePoint = token.startsWith('#x')
      ? Number.parseInt(token.slice(2), 16)
      : Number.parseInt(token.slice(1), 10)
    return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity
  })
}

export function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function readTag(xml: string, start: number): Tag | undefined {
  if (xml[start] !== '<') return undefined
  let index = start + 1
  const marker = xml[index]
  if (marker === '?' || marker === '!') return undefined
  let kind: Tag['kind'] = 'open'
  if (marker === '/') {
    kind = 'close'
    index += 1
  }
  const nameStart = index
  while (index < xml.length && /[A-Za-z0-9_.:-]/.test(xml[index] ?? '')) index += 1
  const full = xml.slice(nameStart, index)
  if (full.length === 0) return undefined
  const colon = full.indexOf(':')
  const prefix = colon === -1 ? '' : full.slice(0, colon)
  const name = colon === -1 ? full : full.slice(colon + 1)
  const attrsStart = index
  let quote = ''
  while (index < xml.length) {
    const char = xml[index] ?? ''
    if (quote !== '') {
      if (char === quote) quote = ''
      index += 1
      continue
    }
    if (char === '"' || char === "'") {
      quote = char
      index += 1
      continue
    }
    if (char === '>') break
    index += 1
  }
  const empty = kind === 'open' && xml[index - 1] === '/'
  const attrsEnd = empty ? index - 1 : index
  return {
    kind: empty ? 'empty' : kind,
    prefix,
    name,
    attrs: xml.slice(attrsStart, attrsEnd),
    start,
    end: index + 1
  }
}
