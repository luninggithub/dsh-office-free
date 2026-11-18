import {
  EMU_PER_PIXEL,
  openPptx,
  type PaintFill,
  type PaintParagraph,
  type PaintRun,
  type PaintShape,
  type PptxDocument,
  type SlideScene
} from '../src/index.ts'

/**
 * Web shell. The slide is painted from theme colors, master and layout chrome,
 * shape fills, and run fonts. Editable text is changed on the slide itself.
 * Line breaking is still a simple wrap, not Office layout.
 */

const page = {
  doc: null as PptxDocument | null,
  slideIndex: 0,
  scale: 1,
  selectedId: undefined as string | undefined,
  editing: false,
  editorShape: undefined as string | undefined,
  thumbIndex: -1,
  images: new Map<string, ImageBitmap>()
}

let pressing = false

const openButton = document.querySelector<HTMLButtonElement>('#open')
const fileInput = document.querySelector<HTMLInputElement>('#file')
const fileName = document.querySelector<HTMLSpanElement>('#file-name')
const saveButton = document.querySelector<HTMLButtonElement>('#save')
const thumbs = document.querySelector<HTMLElement>('#thumbs')
const sheet = document.querySelector<HTMLDivElement>('#sheet')
const canvas = document.querySelector<HTMLCanvasElement>('#slide')
const editorFrame = document.querySelector<HTMLDivElement>('#editor-frame')
const editor = document.querySelector<HTMLDivElement>('#editor')
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
sheet?.addEventListener('pointerdown', (event) => {
  if (event.button !== 0 || canvas === null || !(event.target instanceof Node)) return
  pressing = true
  const box = canvas.getBoundingClientRect()
  const shape = hitTest(event.clientX - box.left, event.clientY - box.top)
  const sameText =
    page.editing &&
    shape?.id === page.selectedId &&
    editor !== null &&
    (event.target === editor || editor.contains(event.target))
  if (sameText) return
  // Keep the pointer from moving focus. A blur while the pointer is down hides the editor and drops the click.
  event.preventDefault()
  page.selectedId = shape?.id
  page.editing = shape?.editable === true
  render()
  if (shape?.editable === true) placeCaret(event.clientX, event.clientY)
})
sheet?.addEventListener('mousedown', (event) => {
  if (event.button !== 0 || !(event.target instanceof Node)) return
  if (editor !== null && (event.target === editor || editor.contains(event.target))) return
  event.preventDefault()
})
window.addEventListener('pointerup', () => {
  if (!pressing) return
  pressing = false
  if (page.editing && editor !== null && document.activeElement !== editor) editor.focus()
})
window.addEventListener('pointercancel', () => {
  pressing = false
})
canvas?.addEventListener('mousemove', (event) => {
  const shape = hitTest(event.offsetX, event.offsetY)
  canvas.style.cursor = shape?.editable === true ? 'text' : 'default'
})
editor?.addEventListener('input', () => {
  const doc = page.doc
  const shape = selected()
  if (doc === null || editor === null || shape === undefined || !shape.editable) return
  try {
    doc.setShapeText(page.slideIndex, shape.id, editor.textContent ?? '')
  } catch (error) {
    note(error instanceof Error ? error.message : '无法修改文字')
    return
  }
  render()
})
editor?.addEventListener('beforeinput', (event) => {
  if (event.inputType !== 'insertParagraph' && event.inputType !== 'insertLineBreak') return
  event.preventDefault()
  insertPlain('\n')
})
editor?.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') return
  event.preventDefault()
  page.editing = false
  render()
  editor.blur()
})
editor?.addEventListener('blur', () => {
  window.setTimeout(() => {
    if (pressing || document.activeElement === editor) return
    page.editing = false
    render()
  }, 0)
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
  page.editing = false
  page.editorShape = undefined
  page.thumbIndex = -1
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

function syncEditor(shape: PaintShape | undefined): void {
  if (editorFrame === null || editor === null) return
  if (shape === undefined || !shape.editable || !page.editing) {
    editorFrame.hidden = true
    page.editorShape = undefined
    return
  }
  const switched = page.editorShape !== shape.id
  page.editorShape = shape.id
  editorFrame.hidden = false
  placeEditor(shape)
  if (document.activeElement !== editor || switched) editor.textContent = shape.plain
}

function placeEditor(shape: PaintShape): void {
  if (editorFrame === null || editor === null) return
  const scale = page.scale
  const text = shape.text
  const run = text?.paragraphs[0]?.runs[0]
  const paragraph = text?.paragraphs[0]
  editorFrame.style.left = `${(shape.x / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.top = `${(shape.y / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.width = `${(shape.width / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.height = `${(shape.height / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.transform = `rotate(${shape.rotation}deg)`
  const pad = text?.padding
  editorFrame.style.padding =
    pad === undefined
      ? '0'
      : `${(pad.top / EMU_PER_PIXEL) * scale}px ${(pad.right / EMU_PER_PIXEL) * scale}px ${(pad.bottom / EMU_PER_PIXEL) * scale}px ${(pad.left / EMU_PER_PIXEL) * scale}px`
  const anchor = text?.anchor ?? 't'
  editorFrame.style.justifyContent =
    anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start'
  const fontPx = ((run?.fontSize ?? 18) * 96) / 72
  editor.style.fontSize = `${fontPx * scale}px`
  editor.style.fontWeight = run?.bold === true ? '700' : '400'
  editor.style.fontStyle = run?.italic === true ? 'italic' : ''
  editor.style.fontFamily =
    run === undefined ? 'Calibri, "Microsoft YaHei", sans-serif' : editorFamily(run)
  editor.style.color = run?.color ?? '#000000'
  editor.style.textAlign =
    paragraph?.align === 'ctr' ? 'center' : paragraph?.align === 'r' ? 'right' : 'left'
  editor.style.textDecoration = run?.underline === true ? 'underline' : 'none'
  const line = paragraph === undefined ? fontPx * 1.2 : lineBox(run, paragraph)
  editor.style.lineHeight = `${line * scale}px`
}

function editorFamily(run: PaintRun): string {
  return [
    quotedFace(run.eastAsian),
    quotedFace(run.latin),
    '"Microsoft YaHei"',
    '"微软雅黑"',
    'Calibri',
    'Arial',
    'sans-serif'
  ]
    .filter((item) => item.length > 0)
    .join(', ')
}

function quotedFace(name: string): string {
  return name.length > 0 ? `"${name.replace(/"/g, '')}"` : ''
}

function placeCaret(clientX: number, clientY: number): void {
  if (editor === null) return
  editor.focus()
  void editor.offsetWidth
  const range = caretRange(clientX, clientY)
  const selection = document.getSelection()
  if (range !== null && editor.contains(range.startContainer)) {
    range.collapse(true)
    selection?.removeAllRanges()
    selection?.addRange(range)
    return
  }
  const end = document.createRange()
  end.selectNodeContents(editor)
  end.collapse(false)
  selection?.removeAllRanges()
  selection?.addRange(end)
}

function caretRange(clientX: number, clientY: number): Range | null {
  if (document.caretRangeFromPoint !== undefined)
    return document.caretRangeFromPoint(clientX, clientY)
  const positioned = document.caretPositionFromPoint?.(clientX, clientY)
  if (positioned === null || positioned === undefined) return null
  const range = document.createRange()
  range.setStart(positioned.offsetNode, positioned.offset)
  range.collapse(true)
  return range
}

function insertPlain(value: string): void {
  if (editor === null) return
  const selection = document.getSelection()
  if (selection === null || selection.rangeCount === 0 || !editor.contains(selection.anchorNode))
    return
  const range = selection.getRangeAt(0)
  range.deleteContents()
  const node = document.createTextNode(value)
  range.insertNode(node)
  range.setStartAfter(node)
  range.collapse(true)
  selection.removeAllRanges()
  selection.addRange(range)
  editor.dispatchEvent(new Event('input', { bubbles: true }))
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
  if (canvas === null || context === null || context === undefined) return
  const doc = page.doc
  if (doc === null) {
    thumbs?.replaceChildren()
    syncEditor(undefined)
    return
  }
  syncThumbs(doc)
  const scene = doc.scene(page.slideIndex)
  const width = Math.max(1, scene.width / EMU_PER_PIXEL)
  const height = Math.max(1, scene.height / EMU_PER_PIXEL)
  const stage = canvas.closest('.stage')
  const available = Math.max(1, (stage?.clientWidth ?? width + 48) - 48)
  const scale = Math.min(1, available / width)
  page.scale = scale
  const ratio = window.devicePixelRatio || 1
  paintCanvas(
    canvas,
    context,
    scene,
    width,
    height,
    scale,
    ratio,
    page.editing ? page.selectedId : undefined
  )
  const shape = selected()
  if (shape !== undefined && shape.selectable) strokeSelection(context, shape)
  syncEditor(shape)
}

function syncThumbs(doc: PptxDocument): void {
  if (thumbs === null) return
  const slides = doc.slides()
  while (thumbs.children.length > slides.length) thumbs.lastElementChild?.remove()
  const slideChanged = page.thumbIndex !== page.slideIndex
  slides.forEach((_, index) => {
    const existing = thumbs.children.item(index)
    const button = existing instanceof HTMLButtonElement ? existing : createThumb()
    if (existing !== button) thumbs.append(button)
    button.dataset.index = String(index)
    button.classList.toggle('current', index === page.slideIndex)
    if (index === page.slideIndex) button.setAttribute('aria-current', 'true')
    else button.removeAttribute('aria-current')
    const label = button.querySelector('.thumb-index')
    if (label !== null) label.textContent = String(index + 1)
    const preview = button.querySelector('canvas')
    if (preview !== null) paintThumb(preview, doc.scene(index))
    if (slideChanged && index === page.slideIndex) button.scrollIntoView({ block: 'nearest' })
  })
  page.thumbIndex = page.slideIndex
}

function createThumb(): HTMLButtonElement {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = 'thumb'
  const label = document.createElement('span')
  label.className = 'thumb-index'
  button.append(label, document.createElement('canvas'))
  button.addEventListener('pointerdown', (event: PointerEvent) => {
    if (event.button !== 0) return
    event.preventDefault()
    const next = Number(button.dataset.index)
    if (!Number.isInteger(next)) return
    page.slideIndex = next
    page.selectedId = undefined
    page.editing = false
    if (document.activeElement === editor) editor?.blur()
    render()
  })
  return button
}

function paintThumb(preview: HTMLCanvasElement, scene: SlideScene): void {
  const ctx = preview.getContext('2d')
  if (ctx === null) return
  const width = Math.max(1, scene.width / EMU_PER_PIXEL)
  const height = Math.max(1, scene.height / EMU_PER_PIXEL)
  const cssWidth = 148
  const scale = cssWidth / width
  const ratio = window.devicePixelRatio || 1
  paintCanvas(preview, ctx, scene, width, height, scale, ratio, undefined)
}

function paintCanvas(
  target: HTMLCanvasElement,
  ctx: CanvasRenderingContext2D,
  scene: SlideScene,
  width: number,
  height: number,
  scale: number,
  ratio: number,
  omitTextId: string | undefined
): void {
  target.width = Math.max(1, Math.round(width * scale * ratio))
  target.height = Math.max(1, Math.round(height * scale * ratio))
  target.style.width = `${width * scale}px`
  target.style.height = `${height * scale}px`
  ctx.setTransform(ratio * scale, 0, 0, ratio * scale, 0, 0)
  ctx.clearRect(0, 0, width, height)
  ctx.fillStyle = '#ffffff'
  ctx.fillRect(0, 0, width, height)
  paintBackdrop(ctx, scene.background, width, height)
  for (const shape of scene.shapes) drawShape(ctx, shape, omitTextId)
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

function drawShape(
  ctx: CanvasRenderingContext2D,
  shape: PaintShape,
  omitTextId: string | undefined
): void {
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
  if (shape.text !== undefined && shape.id !== omitTextId) drawText(ctx, shape, x, y, width, height)
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
  readonly ascent: number
  readonly descent: number
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
        const measured = ctx.measureText(token.text)
        const width = measured.width
        if (width > contentWidth && token.text.length > 1) {
          for (const char of token.text) {
            const charMeasured = ctx.measureText(char)
            const charWidth = charMeasured.width
            if (line.width + charWidth > contentWidth && line.width > 0) push()
            addToken(line, char, charWidth, charMeasured, token.cjk, run, paragraph)
          }
          continue
        }
        if (line.width + width > contentWidth && line.width > 0) push()
        addToken(line, token.text, width, measured, token.cjk, run, paragraph)
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
  measured: TextMetrics,
  cjk: boolean,
  run: PaintRun,
  paragraph: PaintParagraph
): void {
  const size = (run.fontSize * 96) / 72
  const { ascent, descent } = fontEdges(measured, size)
  line.tokens.push({ text, width, ascent, descent, cjk, run })
  line.width += width
  line.height = Math.max(line.height, lineBox(run, paragraph))
}

function fontEdges(measured: TextMetrics, size: number): { ascent: number; descent: number } {
  const ascent = measured.fontBoundingBoxAscent
  const descent = measured.fontBoundingBoxDescent
  if (ascent > 0 || descent > 0) return { ascent, descent }
  return { ascent: size * 0.8, descent: size * 0.2 }
}

function lineBaseline(top: number, height: number, ascent: number, descent: number): number {
  return top + (height - (ascent + descent)) / 2 + ascent
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
    ctx.textBaseline = 'alphabetic'
    if (index === 0 && paragraph.bullet !== undefined && paragraph.bulletRun !== undefined) {
      ctx.font = fontOf(paragraph.bulletRun, isCjk(paragraph.bullet))
      ctx.fillStyle = paragraph.bulletRun.color
      const size = (paragraph.bulletRun.fontSize * 96) / 72
      const edge = fontEdges(ctx.measureText(paragraph.bullet), size)
      ctx.fillText(
        paragraph.bullet,
        textX + paragraph.indent,
        lineBaseline(cursor, line.height, edge.ascent, edge.descent)
      )
    }
    for (const token of line.tokens) {
      ctx.font = fontOf(token.run, token.cjk)
      ctx.fillStyle = token.run.color
      const baseline = lineBaseline(cursor, line.height, token.ascent, token.descent)
      ctx.fillText(token.text, offset, baseline)
      if (token.run.underline) {
        const size = (token.run.fontSize * 96) / 72
        const underline = baseline + size * 0.12
        ctx.strokeStyle = token.run.color
        ctx.lineWidth = Math.max(1, size / 14)
        ctx.beginPath()
        ctx.moveTo(offset, underline)
        ctx.lineTo(offset + token.width, underline)
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
