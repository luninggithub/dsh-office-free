/** ZIP reader and writer for OOXML packages. Methods are store (0) and deflate (8). */

export interface ZipEntry {
  readonly name: string
  readonly data: Uint8Array
  /** Defaults to store, which OOXML readers accept. */
  readonly method?: 'store' | 'deflate'
}

const textEncoder = new TextEncoder()

/** Read every file in a ZIP archive. Directory entries are omitted. */
export async function readZip(archive: Uint8Array): Promise<Map<string, Uint8Array>> {
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength)
  const end = findEndOfCentralDirectory(archive)
  const count = view.getUint16(end + 8, true)
  let cursor = view.getUint32(end + 16, true)
  const pending: Array<{ name: string; data: Promise<Uint8Array> }> = []
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(cursor, true) !== 0x02014b50)
      throw new Error('pptx: malformed ZIP directory')
    const method = view.getUint16(cursor + 10, true)
    const compressedSize = view.getUint32(cursor + 20, true)
    const nameLength = view.getUint16(cursor + 28, true)
    const extraLength = view.getUint16(cursor + 30, true)
    const commentLength = view.getUint16(cursor + 32, true)
    const localOffset = view.getUint32(cursor + 42, true)
    const name = new TextDecoder().decode(archive.subarray(cursor + 46, cursor + 46 + nameLength))
    const nameLengthLocal = view.getUint16(localOffset + 26, true)
    const extraLengthLocal = view.getUint16(localOffset + 28, true)
    const dataStart = localOffset + 30 + nameLengthLocal + extraLengthLocal
    const compressed = archive.subarray(dataStart, dataStart + compressedSize)
    if (!name.endsWith('/')) {
      if (method === 0) pending.push({ name, data: Promise.resolve(compressed.slice()) })
      else if (method === 8) pending.push({ name, data: inflateRaw(compressed) })
      else throw new Error(`pptx: unsupported ZIP method ${method} for ${name}`)
    }
    cursor += 46 + nameLength + extraLength + commentLength
  }
  const entries = new Map<string, Uint8Array>()
  const inflated = await Promise.all(pending.map(async (item) => item.data))
  pending.forEach((item, index) => {
    const data = inflated[index]
    if (data !== undefined) entries.set(item.name, data)
  })
  return entries
}

/** Write a ZIP archive. Entry order is preserved so later saves stay stable. */
export async function writeZip(entries: readonly ZipEntry[]): Promise<Uint8Array> {
  const prepared = await Promise.all(
    entries.map(async (entry) => {
      const method = entry.method === 'deflate' ? 8 : 0
      return {
        nameBytes: textEncoder.encode(entry.name),
        method,
        stored: entry.data,
        data: method === 8 ? await deflateRaw(entry.data) : entry.data,
        crc: crc32(entry.data)
      }
    })
  )
  const locals: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  for (const entry of prepared) {
    const { nameBytes, method, stored, data, crc } = entry
    const local = new Uint8Array(30)
    const localView = new DataView(local.buffer)
    localView.setUint32(0, 0x04034b50, true)
    localView.setUint16(4, 20, true)
    localView.setUint16(8, method, true)
    localView.setUint32(14, crc, true)
    localView.setUint32(18, data.byteLength, true)
    localView.setUint32(22, stored.byteLength, true)
    localView.setUint16(26, nameBytes.byteLength, true)
    locals.push(local, nameBytes, data)
    const directory = new Uint8Array(46)
    const directoryView = new DataView(directory.buffer)
    directoryView.setUint32(0, 0x02014b50, true)
    directoryView.setUint16(4, 20, true)
    directoryView.setUint16(6, 20, true)
    directoryView.setUint16(10, method, true)
    directoryView.setUint32(16, crc, true)
    directoryView.setUint32(20, data.byteLength, true)
    directoryView.setUint32(24, stored.byteLength, true)
    directoryView.setUint16(28, nameBytes.byteLength, true)
    directoryView.setUint32(42, offset, true)
    central.push(directory, nameBytes)
    offset += local.byteLength + nameBytes.byteLength + data.byteLength
  }
  const centralBytes = concat(central)
  const end = new Uint8Array(22)
  const endView = new DataView(end.buffer)
  endView.setUint32(0, 0x06054b50, true)
  endView.setUint16(8, entries.length, true)
  endView.setUint16(10, entries.length, true)
  endView.setUint32(12, centralBytes.byteLength, true)
  endView.setUint32(16, offset, true)
  return concat([...locals, centralBytes, end])
}

function findEndOfCentralDirectory(archive: Uint8Array): number {
  const min = Math.max(0, archive.length - 22 - 0xffff)
  for (let index = archive.length - 22; index >= min; index -= 1) {
    if (
      archive[index] === 0x50 &&
      archive[index + 1] === 0x4b &&
      archive[index + 2] === 0x05 &&
      archive[index + 3] === 0x06
    ) {
      return index
    }
  }
  throw new Error('pptx: ZIP end of central directory not found')
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([toBody(data)])
    .stream()
    .pipeThrough(new DecompressionStream('deflate-raw'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

async function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([toBody(data)]).stream().pipeThrough(new CompressionStream('deflate-raw'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

/** Blob accepts a BufferSource; a fresh ArrayBuffer view avoids a SharedArrayBuffer typing mismatch. */
function toBody(data: Uint8Array): ArrayBuffer {
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const length = parts.reduce((sum, part) => sum + part.byteLength, 0)
  const out = new Uint8Array(length)
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.byteLength
  }
  return out
}

let crcTable: Uint32Array | undefined

function crc32(bytes: Uint8Array): number {
  if (crcTable === undefined) {
    crcTable = new Uint32Array(256)
    for (let n = 0; n < 256; n += 1) {
      let c = n
      for (let k = 0; k < 8; k += 1) c = (c & 1) === 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      crcTable[n] = c
    }
  }
  let crc = 0xffffffff
  for (const byte of bytes) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff]!
  return (crc ^ 0xffffffff) >>> 0
}
