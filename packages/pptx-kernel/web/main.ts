import {
  EMU_PER_PIXEL,
  openPptx,
  type PaintFill,
  type PaintParagraph,
  type PaintRun,
  type PaintShape,
  type PptxDocument
} from '../src/index.ts'

/**
 * Web shell. The slide is painted from theme colors, master and layout chrome,
 * shape fills, and run fonts. Line breaking is still a simple wrap, not Office layout.
 */

const page = {
  doc: null as PptxDocument | null,
  slideIndex: 0,
  scale: 1,
  selectedId: undefined as string | undefined,
  images: new Map<string, ImageBitmap>()
}

const openButton = document.querySelector<HTMLButtonElement>('#open')
const fileInput = document.querySelector<HTMLInputElement>('#file')
const fileName = document.querySelector<HTMLSpanElement>('#file-name')
const saveButton = document.querySelector<HTMLButtonElement>('#save')
const pages = document.querySelector<HTMLDivElement>('#pages')
const canvas = document.querySelector<HTMLCanvasElement>('#slide')
const shapeName = document.querySelector<HTMLDivElement>('#shape-name')
const textInput = document.querySelector<HTMLTextAreaElement>('#text')
const context = canvas?.getContext('2d')

openButton?.addEventListener('click', () => fileInput?.click())
fileInput?.addEventListener('change', () => {
  const file = fileInput.files?.[0]
  if (file !== undefined) void openFile(file)
})
saveButton?.addEventListener('click', () => {
  const doc = page.doc
  if (doc === null) return
  void doc.save().then((bytes) => download(bytes))
})
canvas?.addEventListener('click', (event) => {
  const shape = hitTest(event.offsetX, event.offsetY)
  page.selectedId = shape?.id
  render()
})
textInput?.addEventListener('input', () => {
  const doc = page.doc
  const shape = selected()
  if (doc === null || shape === undefined || !shape.editable || textInput === null) return
  doc.setShapeText(page.slideIndex, shape.id, textInput.value)
  render()
})
window.addEventListener('resize', () => {
  if (page.doc !== null) render()
})
document.body.addEventListener('dragover', (event) => {
  event.preventDefault()
})
document.body.addEventListener('drop', (event) => {
  event.preventDefault()
  const file = event.dataTransfer?.files[0]
  if (file !== undefined) void openFile(file)
})

const requested = new URLSearchParams(location.search).get('open')
if (requested !== null && safeFileName(requested)) void openRequested(requested)

async function openRequested(name: string): Promise<void> {
  try {
    const response = await fetch(encodeURIComponent(name))
    if (!response.ok) throw new Error('pptx: 找不到这个文件')
    await openFile(new File([await response.arrayBuffer()], name))
  } catch (error) {
    note(error instanceof Error ? error.message : '无法打开这个文件')
  }
}

async function openFile(file: File): Promise<void> {
  try {
    page.doc = await openPptx(new Uint8Array(await file.arrayBuffer()))
  } catch (error) {
    page.doc = null
    if (saveButton !== null) saveButton.disabled = true
    note(error instanceof Error ? error.message : '无法打开这个文件')
    render()
    return
  }
  page.slideIndex = 0
  page.selectedId = undefined
  page.images.clear()
  if (saveButton !== null) saveButton.disabled = false
  note(file.name)
  await loadImages()
  render()
}

function note(message: string): void {
  if (fileName !== null) fileName.textContent = message
}

function safeFileName(name: string): boolean {
  return name.length > 0 && !name.includes('/') && !name.includes('\\') && !name.includes('..')
}

function selected(): PaintShape | undefined {
  return page.doc?.scene(page.slideIndex).shapes.find((shape) => shape.id === page.selectedId)
}

function hitTest(x: number, y: number): PaintShape | undefined {
  const shapes = page.doc?.scene(page.slideIndex).shapes ?? []
  const px = x / page.scale
  const py = y / page.scale
  for (let index = shapes.length - 1; index >= 0; index -= 1) {
    const shape = shapes[index]
    if (shape === undefined || !shape.selectable) continue
    if (hits(shape, px, py)) return shape
  }
  return undefined
}

function hits(shape: PaintShape, px: number, py: number): boolean {
  const x = shape.x / EMU_PER_PIXEL
  const y = shape.y / EMU_PER_PIXEL
  const width = shape.width / EMU_PER_PIXEL
  const height = shape.height / EMU_PER_PIXEL
  const dx = px - (x + width / 2)
  const dy = py - (y + height / 2)
  const rad = (shape.rotation * Math.PI) / 180
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)
  const localX = dx * cos + dy * sin
  const localY = -dx * sin + dy * cos
  return Math.abs(localX) <= width / 2 && Math.abs(localY) <= height / 2
}

async function loadImages(): Promise<void> {
  const doc = page.doc
  if (doc === null) return
  const paths = new Set<string>()
  doc.slides().forEach((_, index) => {
    for (const path of doc.scene(index).imagePaths) paths.add(path)
  })
  await Promise.all(
    [...paths].map(async (path) => {
      if (page.images.has(path)) return
      const bytes = doc.part(path)
      if (bytes === undefined) return
      try {
        page.images.set(path, await createImageBitmap(new Blob([toBuffer(bytes)])))
      } catch {
        // Browser image decoders do not accept EMF/WMF.
      }
    })
  )
}

function render(): void {
  if (
    canvas === null ||
    context === null ||
    context === undefined ||
    pages === null ||
    shapeName === null ||
    textInput === null
  ) {
    return
  }
  const doc = page.doc
  pages.replaceChildren()
  if (doc === null) return
  const slides = doc.slides()
  slides.forEach((_slide, index) => {
    const button = document.createElement('button')
    button.type = 'button'
    button.textContent = `第 ${index + 1} 页`
    if (index === page.slideIndex) button.className = 'current'
    button.addEventListener('click', () => {
      page.slideIndex = index
      page.selectedId = undefined
      render()
    })
    pages.append(button)
  })
  const scene = doc.scene(page.slideIndex)
  const width = Math.max(1, scene.width / EMU_PER_PIXEL)
  const height = Math.max(1, scene.height / EMU_PER_PIXEL)
  const available = Math.max(1, (canvas.parentElement?.clientWidth ?? width) - 48)
  const scale = Math.min(1, available / width)
  page.scale = scale
  const ratio = window.devicePixelRatio || 1
  canvas.width = Math.round(width * scale * ratio)
  canvas.height = Math.round(height * scale * ratio)
  canvas.style.width = `${width * scale}px`
  canvas.style.height = `${height * scale}px`
  context.setTransform(ratio * scale, 0, 0, ratio * scale, 0, 0)
  context.clearRect(0, 0, width, height)
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, width, height)
  paintBackdrop(context, scene.background, width, height)
  for (const shape of scene.shapes) drawShape(context, shape)
  const shape = selected()
  if (shape !== undefined && shape.selectable) strokeSelection(context, shape)
  if (shape === undefined) {
    shapeName.textContent = '未选中形状'
    textInput.value = ''
    textInput.disabled = true
    return
  }
  shapeName.textContent = shape.name.length > 0 ? shape.name : `形状 ${shape.id}`
  textInput.disabled = !shape.editable
  if (document.activeElement !== textInput) textInput.value = shape.plain
}

function paintBackdrop(
  ctx: CanvasRenderingContext2D,
  fill: PaintFill,
  width: number,
  height: number
): void {
  if (fill.kind === 'none') return
  if (fill.kind === 'picture') {
    const bitmap = page.images.get(fill.imagePath)
    if (bitmap !== undefined) drawBitmap(ctx, bitmap, fill.crop, 0, 0, width, height)
    return
  }
  ctx.fillStyle = paintFillStyle(ctx, fill, 0, 0, width, height)
  ctx.fillRect(0, 0, width, height)
}

function drawShape(ctx: CanvasRenderingContext2D, shape: PaintShape): void {
  const x = shape.x / EMU_PER_PIXEL
  const y = shape.y / EMU_PER_PIXEL
  const width = shape.width / EMU_PER_PIXEL
  const height = shape.height / EMU_PER_PIXEL
  ctx.save()
  const cx = x + width / 2
  const cy = y + height / 2
  ctx.translate(cx, cy)
  ctx.rotate((shape.rotation * Math.PI) / 180)
  if (shape.flipH || shape.flipV) ctx.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1)
  ctx.translate(-cx, -cy)
  if (shape.geometry === 'line') {
    strokeLine(ctx, shape, x, y, width, height)
    ctx.restore()
    return
  }
  if (shape.shadow !== undefined) {
    ctx.shadowColor = shape.shadow.color
    ctx.shadowBlur = shape.shadow.blur / EMU_PER_PIXEL
    ctx.shadowOffsetX = shape.shadow.offsetX / EMU_PER_PIXEL
    ctx.shadowOffsetY = shape.shadow.offsetY / EMU_PER_PIXEL
  }
  trace(ctx, shape, x, y, width, height)
  if (shape.fill.kind === 'solid' || shape.fill.kind === 'linear' || shape.fill.kind === 'radial') {
    ctx.fillStyle = paintFillStyle(ctx, shape.fill, x, y, width, height)
    ctx.fill()
  } else if (shape.shadow !== undefined && shape.fill.kind === 'picture') {
    ctx.fillStyle = '#ffffff'
    ctx.fill()
  }
  ctx.shadowColor = 'transparent'
  if (shape.fill.kind === 'picture') {
    ctx.save()
    trace(ctx, shape, x, y, width, height)
    ctx.clip()
    const bitmap = page.images.get(shape.fill.imagePath)
    if (bitmap === undefined) {
      ctx.fillStyle = '#e6e8ea'
      ctx.fillRect(x, y, width, height)
    } else {
      drawBitmap(ctx, bitmap, shape.fill.crop, x, y, width, height)
    }
    ctx.restore()
  }
  if (shape.line !== undefined) {
    ctx.strokeStyle = shape.line.color
    ctx.lineWidth = Math.max(shape.line.width / EMU_PER_PIXEL, 0.5)
    setDash(ctx, shape.line.dash)
    trace(ctx, shape, x, y, width, height)
    ctx.stroke()
    ctx.setLineDash([])
  }
  if (shape.text !== undefined) drawText(ctx, shape, x, y, width, height)
  ctx.restore()
}

function strokeSelection(ctx: CanvasRenderingContext2D, shape: PaintShape): void {
  const x = shape.x / EMU_PER_PIXEL
  const y = shape.y / EMU_PER_PIXEL
  const width = shape.width / EMU_PER_PIXEL
  const height = shape.height / EMU_PER_PIXEL
  ctx.save()
  const cx = x + width / 2
  const cy = y + height / 2
  ctx.translate(cx, cy)
  ctx.rotate((shape.rotation * Math.PI) / 180)
  ctx.translate(-cx, -cy)
  ctx.strokeStyle = '#14a6ea'
  ctx.lineWidth = 2 / page.scale
  ctx.setLineDash([])
  ctx.strokeRect(x, y, width, height)
  ctx.restore()
}

function strokeLine(
  ctx: CanvasRenderingContext2D,
  shape: PaintShape,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  ctx.strokeStyle = shape.line?.color ?? '#000000'
  ctx.lineWidth = Math.max((shape.line?.width ?? 9525) / EMU_PER_PIXEL, 0.75)
  setDash(ctx, shape.line?.dash ?? 'solid')
  ctx.beginPath()
  ctx.moveTo(x, y)
  ctx.lineTo(x + width, y + height)
  ctx.stroke()
  ctx.setLineDash([])
}

function trace(
  ctx: CanvasRenderingContext2D,
  shape: PaintShape,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  ctx.beginPath()
  if (shape.geometry === 'ellipse') {
    ctx.ellipse(
      x + width / 2,
      y + height / 2,
      Math.abs(width) / 2,
      Math.abs(height) / 2,
      0,
      0,
      Math.PI * 2
    )
    return
  }
  if (shape.geometry === 'round') {
    const radius = Math.min(Math.abs(width), Math.abs(height)) * shape.corner
    ctx.roundRect(x, y, width, height, radius)
    return
  }
  if (shape.geometry === 'triangle') {
    ctx.moveTo(x + width / 2, y)
    ctx.lineTo(x + width, y + height)
    ctx.lineTo(x, y + height)
    ctx.closePath()
    return
  }
  if (shape.geometry === 'diamond') {
    ctx.moveTo(x + width / 2, y)
    ctx.lineTo(x + width, y + height / 2)
    ctx.lineTo(x + width / 2, y + height)
    ctx.lineTo(x, y + height / 2)
    ctx.closePath()
    return
  }
  ctx.rect(x, y, width, height)
}

function paintFillStyle(
  ctx: CanvasRenderingContext2D,
  fill: Exclude<PaintFill, { kind: 'none' } | { kind: 'picture' }>,
  x: number,
  y: number,
  width: number,
  height: number
): string | CanvasGradient {
  if (fill.kind === 'solid') return fill.color
  if (fill.kind === 'radial') {
    const radius = Math.max(width, height) / 2
    const gradient = ctx.createRadialGradient(
      x + width / 2,
      y + height / 2,
      0,
      x + width / 2,
      y + height / 2,
      Math.max(1, radius)
    )
    addStops(gradient, fill.stops)
    return gradient
  }
  let rad = (fill.angle * Math.PI) / 180
  if (fill.scaled && width !== 0 && height !== 0) {
    rad = Math.atan2(Math.sin(rad) * height, Math.cos(rad) * width)
  }
  const dx = Math.cos(rad)
  const dy = Math.sin(rad)
  const cx = x + width / 2
  const cy = y + height / 2
  const half = Math.abs((width / 2) * dx) + Math.abs((height / 2) * dy)
  const gradient = ctx.createLinearGradient(
    cx - dx * half,
    cy - dy * half,
    cx + dx * half,
    cy + dy * half
  )
  addStops(gradient, fill.stops)
  return gradient
}

function addStops(
  gradient: CanvasGradient,
  stops: readonly { offset: number; color: string }[]
): void {
  const ordered = [...stops].toSorted((a, b) => a.offset - b.offset)
  let last = -1
  for (const stop of ordered) {
    let offset = Math.min(1, Math.max(0, stop.offset))
    if (offset <= last) offset = Math.min(1, last + 0.0001)
    gradient.addColorStop(offset, stop.color)
    last = offset
  }
  if (ordered.length === 0) gradient.addColorStop(0, '#ffffff')
}

function drawBitmap(
  ctx: CanvasRenderingContext2D,
  bitmap: ImageBitmap,
  crop: { left: number; top: number; right: number; bottom: number } | undefined,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  const left = crop?.left ?? 0
  const top = crop?.top ?? 0
  const right = crop?.right ?? 0
  const bottom = crop?.bottom ?? 0
  const sx = bitmap.width * left
  const sy = bitmap.height * top
  const sw = bitmap.width * (1 - left - right)
  const sh = bitmap.height * (1 - top - bottom)
  if (sw <= 0 || sh <= 0) return
  ctx.drawImage(bitmap, sx, sy, sw, sh, x, y, width, height)
}

function drawText(
  ctx: CanvasRenderingContext2D,
  shape: PaintShape,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  const text = shape.text
  if (text === undefined) return
  const pad = {
    left: text.padding.left / EMU_PER_PIXEL,
    top: text.padding.top / EMU_PER_PIXEL,
    right: text.padding.right / EMU_PER_PIXEL,
    bottom: text.padding.bottom / EMU_PER_PIXEL
  }
  const innerWidth = Math.max(1, width - pad.left - pad.right)
  const innerHeight = Math.max(0, height - pad.top - pad.bottom)
  ctx.save()
  trace(ctx, shape, x, y, width, height)
  ctx.clip()
  ctx.textBaseline = 'top'
  const laid = text.paragraphs.map((paragraph) => layoutParagraph(ctx, paragraph, innerWidth))
  const contentHeight = laid.reduce((sum, paragraph) => sum + paragraph.height, 0)
  let cursor = y + pad.top
  if (text.anchor === 'ctr') cursor += Math.max(0, (innerHeight - contentHeight) / 2)
  if (text.anchor === 'b') cursor += Math.max(0, innerHeight - contentHeight)
  for (const paragraph of laid) {
    drawParagraph(ctx, paragraph, x + pad.left, cursor, innerWidth)
    cursor += paragraph.height
  }
  ctx.restore()
}

interface LaidToken {
  readonly text: string
  readonly width: number
  readonly cjk: boolean
  readonly run: PaintRun
}

interface LaidLine {
  width: number
  height: number
  readonly tokens: LaidToken[]
}

interface LaidParagraph {
  readonly align: 'l' | 'ctr' | 'r'
  readonly bullet: string | undefined
  readonly bulletRun: PaintRun | undefined
  readonly marginLeft: number
  readonly indent: number
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly lines: readonly LaidLine[]
  readonly height: number
}

function layoutParagraph(
  ctx: CanvasRenderingContext2D,
  paragraph: PaintParagraph,
  innerWidth: number
): LaidParagraph {
  const marginLeft = paragraph.marginLeft / EMU_PER_PIXEL
  const indent = paragraph.indent / EMU_PER_PIXEL
  const contentWidth = Math.max(1, innerWidth - marginLeft)
  const lines: LaidLine[] = []
  let line = emptyLine(paragraph)
  const push = () => {
    lines.push(line)
    line = emptyLine(paragraph)
  }
  for (const run of paragraph.runs) {
    const parts = run.text.split('\n')
    parts.forEach((part, index) => {
      if (index > 0) push()
      for (const token of tokenize(part)) {
        if (line.tokens.length === 0 && token.text.trim().length === 0) continue
        ctx.font = fontOf(run, token.cjk)
        const width = ctx.measureText(token.text).width
        if (width > contentWidth && token.text.length > 1) {
          for (const char of token.text) {
            const charWidth = ctx.measureText(char).width
            if (line.width + charWidth > contentWidth && line.width > 0) push()
            addToken(line, char, charWidth, token.cjk, run, paragraph)
          }
          continue
        }
        if (line.width + width > contentWidth && line.width > 0) push()
        addToken(line, token.text, width, token.cjk, run, paragraph)
      }
    })
  }
  if (line.tokens.length > 0 || lines.length === 0) lines.push(line)
  const spaceBefore = (paragraph.spaceBefore * 96) / 72
  const spaceAfter = (paragraph.spaceAfter * 96) / 72
  const height = spaceBefore + spaceAfter + lines.reduce((sum, item) => sum + item.height, 0)
  return {
    align: paragraph.align,
    bullet: paragraph.bullet,
    bulletRun: paragraph.runs[0],
    marginLeft,
    indent,
    spaceBefore,
    spaceAfter,
    lines,
    height
  }
}

function emptyLine(paragraph: PaintParagraph): LaidLine {
  return { width: 0, height: lineBox(paragraph.runs[0], paragraph), tokens: [] }
}

function addToken(
  line: LaidLine,
  text: string,
  width: number,
  cjk: boolean,
  run: PaintRun,
  paragraph: PaintParagraph
): void {
  line.tokens.push({ text, width, cjk, run })
  line.width += width
  line.height = Math.max(line.height, lineBox(run, paragraph))
}

function lineBox(run: PaintRun | undefined, paragraph: PaintParagraph): number {
  if (paragraph.linePt > 0) return (paragraph.linePt * 96) / 72
  const size = ((run?.fontSize ?? 18) * 96) / 72
  return size * 1.2 * paragraph.linePct
}

function drawParagraph(
  ctx: CanvasRenderingContext2D,
  paragraph: LaidParagraph,
  x: number,
  y: number,
  innerWidth: number
): void {
  let cursor = y + paragraph.spaceBefore
  const textX = x + paragraph.marginLeft
  paragraph.lines.forEach((line, index) => {
    let offset = textX
    if (paragraph.align === 'ctr') offset = x + (innerWidth - line.width) / 2
    if (paragraph.align === 'r') offset = x + innerWidth - line.width
    if (index === 0 && paragraph.bullet !== undefined && paragraph.bulletRun !== undefined) {
      ctx.font = fontOf(paragraph.bulletRun, isCjk(paragraph.bullet))
      ctx.fillStyle = paragraph.bulletRun.color
      ctx.fillText(paragraph.bullet, textX + paragraph.indent, cursor)
    }
    for (const token of line.tokens) {
      ctx.font = fontOf(token.run, token.cjk)
      ctx.fillStyle = token.run.color
      ctx.fillText(token.text, offset, cursor)
      if (token.run.underline) {
        const size = (token.run.fontSize * 96) / 72
        ctx.strokeStyle = token.run.color
        ctx.lineWidth = Math.max(1, size / 14)
        ctx.beginPath()
        ctx.moveTo(offset, cursor + size * 0.92)
        ctx.lineTo(offset + token.width, cursor + size * 0.92)
        ctx.stroke()
      }
      offset += token.width
    }
    cursor += line.height
  })
}

function fontOf(run: PaintRun, cjk: boolean): string {
  const family = cjk ? run.eastAsian : run.latin
  const size = (run.fontSize * 96) / 72
  const quoted = family.length > 0 ? `"${family.replace(/"/g, '')}"` : ''
  const fallback = cjk
    ? '"Microsoft YaHei", "微软雅黑", "PingFang SC", "Noto Sans SC", sans-serif'
    : 'Calibri, "Segoe UI", Arial, sans-serif'
  const stack = [quoted, fallback].filter((item) => item.length > 0).join(', ')
  return [run.italic ? 'italic' : '', run.bold ? '700' : '400', `${size}px`, stack]
    .filter((item) => item.length > 0)
    .join(' ')
}

function tokenize(text: string): Array<{ text: string; cjk: boolean }> {
  const tokens: Array<{ text: string; cjk: boolean }> = []
  for (const match of text.matchAll(
    /[\u3400-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]|[^\s\u3400-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]+|\s+/gu
  )) {
    const value = match[0] ?? ''
    tokens.push({ text: value, cjk: isCjk(value) })
  }
  return tokens
}

function isCjk(text: string): boolean {
  const code = text.codePointAt(0) ?? 0
  return (
    (code >= 0x3400 && code <= 0x9fff) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0x3040 && code <= 0x30ff) ||
    (code >= 0xac00 && code <= 0xd7af)
  )
}

function setDash(ctx: CanvasRenderingContext2D, dash: 'solid' | 'dash' | 'dot'): void {
  if (dash === 'dot') ctx.setLineDash([1.5, 3])
  else if (dash === 'dash') ctx.setLineDash([8, 4])
  else ctx.setLineDash([])
}

function download(bytes: Uint8Array): void {
  const url = URL.createObjectURL(
    new Blob([toBuffer(bytes)], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    })
  )
  const link = document.createElement('a')
  link.href = url
  link.download = 'edited.pptx'
  link.click()
  URL.revokeObjectURL(url)
}

function toBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}
