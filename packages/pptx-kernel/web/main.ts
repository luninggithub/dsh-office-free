import {
  EMU_PER_PIXEL,
  openPptx,
  type PaintFill,
  type PaintParagraph,
  type PaintRun,
  type PaintShape,
  type PptxCheckpoint,
  type PptxDocument,
  type SlideScene
} from '../src/index.ts'

/**
 * Web shell. The slide is painted from theme colors, master and layout chrome,
 * shape fills, and run fonts. Editable text is changed on the slide itself.
 * Line breaking is still a simple wrap, not Office layout.
 */

interface SavedFile {
  readonly name: string
  getFile(): Promise<File>
  createWritable(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>
}

const page = {
  doc: null as PptxDocument | null,
  slideIndex: 0,
  scale: 1,
  selectedId: undefined as string | undefined,
  editing: false,
  editorShape: undefined as string | undefined,
  thumbIndex: -1,
  images: new Map<string, ImageBitmap>(),
  saved: null as Uint8Array | null,
  dirty: false,
  label: '',
  servedName: null as string | null,
  handle: null as SavedFile | null,
  zoom: 1,
  pageSelected: false,
  cropping: false,
  live: undefined as LiveShape | undefined,
  clip: null as { slide: number; id: string } | null,
  pasteCount: 0
}

interface Box {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

interface LiveShape {
  readonly id: string
  readonly box: Box
  readonly crop?: {
    readonly left: number
    readonly top: number
    readonly right: number
    readonly bottom: number
  }
}

interface ShapeGesture {
  kind: 'pending' | 'move' | 'resize' | 'crop'
  readonly id: string
  readonly pointerId: number
  readonly startClientX: number
  readonly startClientY: number
  readonly origin: Box
  readonly handle: string
  readonly crop: { left: number; top: number; right: number; bottom: number }
  moved: boolean
}

let pressing = false
let painting = false
let applyingHistory = false
let plainOpen = readPlainOpen()
const history: PptxCheckpoint[] = []
let menuIndex = -1
let drag: ThumbDrag | null = null
let shapeGesture: ShapeGesture | null = null

interface ThumbDrag {
  readonly index: number
  readonly pointerId: number
  readonly startY: number
  moved: boolean
}

const openButton = document.querySelector<HTMLButtonElement>('#open')
const insertImageButton = document.querySelector<HTMLButtonElement>('#insert-image')
const fileInput = document.querySelector<HTMLInputElement>('#file')
const imageInput = document.querySelector<HTMLInputElement>('#image-file')
const fileName = document.querySelector<HTMLSpanElement>('#file-name')
const saveButton = document.querySelector<HTMLButtonElement>('#save')
const discardButton = document.querySelector<HTMLButtonElement>('#discard')
const addButton = document.querySelector<HTMLButtonElement>('#add-slide')
const thumbs = document.querySelector<HTMLElement>('#thumbs')
const paneSplit = document.querySelector<HTMLDivElement>('#pane-split')
const slideMenu = document.querySelector<HTMLDivElement>('#slide-menu')
const copySlideButton = document.querySelector<HTMLButtonElement>('#copy-slide')
const deleteSlideButton = document.querySelector<HTMLButtonElement>('#delete-slide')
const viewport = document.querySelector<HTMLDivElement>('#viewport')
const sheet = document.querySelector<HTMLDivElement>('#sheet')
const canvas = document.querySelector<HTMLCanvasElement>('#slide')
const handles = document.querySelector<HTMLDivElement>('#handles')
const cropButton = document.querySelector<HTMLButtonElement>('#crop-shape')
const shapeMenu = document.querySelector<HTMLDivElement>('#shape-menu')
const copyShapeButton = document.querySelector<HTMLButtonElement>('#copy-shape')
const deleteShapeButton = document.querySelector<HTMLButtonElement>('#delete-shape')
const zoomInput = document.querySelector<HTMLInputElement>('#zoom')
const zoomOut = document.querySelector<HTMLButtonElement>('#zoom-out')
const zoomIn = document.querySelector<HTMLButtonElement>('#zoom-in')
const zoomLabel = document.querySelector<HTMLSpanElement>('#zoom-label')
const thumbGhost = document.querySelector<HTMLDivElement>('#thumb-ghost')
const editorFrame = document.querySelector<HTMLDivElement>('#editor-frame')
const editor = document.querySelector<HTMLDivElement>('#editor')
const context = canvas?.getContext('2d')

openButton?.addEventListener('click', () => {
  if (plainOpen || filePicker() === undefined || embeddedFrame()) choosePlainFile()
  else void pickFile()
})
insertImageButton?.addEventListener('click', () => {
  if (imageInput === null) return
  imageInput.value = ''
  imageInput.click()
})
fileInput?.addEventListener('change', () => {
  const file = fileInput.files?.[0]
  fileInput.value = ''
  if (file !== undefined) void openFile(file)
})
imageInput?.addEventListener('change', () => {
  const file = imageInput.files?.[0]
  imageInput.value = ''
  if (file !== undefined) void insertLocalImage(file)
})
saveButton?.addEventListener('click', () => {
  void saveCurrent()
})
discardButton?.addEventListener('click', () => {
  void discardEdits()
})
addButton?.addEventListener('click', () => {
  addSlide()
})
copySlideButton?.addEventListener('click', () => {
  const index = menuIndex
  hideMenu()
  duplicateAt(index)
})
deleteSlideButton?.addEventListener('click', () => {
  const index = menuIndex
  hideMenu()
  deleteAt(index)
})
document.addEventListener('pointerdown', (event) => {
  if (!(event.target instanceof Node)) return
  if (slideMenu !== null && !slideMenu.hidden && !slideMenu.contains(event.target)) hideMenu()
  if (shapeMenu !== null && !shapeMenu.hidden && !shapeMenu.contains(event.target)) hideShapeMenu()
})
window.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    hideMenu()
    hideShapeMenu()
    if (page.cropping) {
      page.cropping = false
      render()
    }
    return
  }
  const command = event.ctrlKey || event.metaKey
  const key = event.key.toLowerCase()
  if (command && key === 'z' && !event.shiftKey && !fieldTarget()) {
    event.preventDefault()
    undo()
    return
  }
  if (typingTarget()) return
  if (event.key === 'Delete' || event.key === 'Backspace') {
    event.preventDefault()
    deleteSelectedShape()
    return
  }
  if (command && key === 'c') {
    if (copySelectedPicture()) event.preventDefault()
    return
  }
  if (command && key === 'v') {
    if (page.clip === null) return
    event.preventDefault()
    pastePicture()
  }
})
copyShapeButton?.addEventListener('click', () => {
  copySelectedPicture()
  hideShapeMenu()
})
deleteShapeButton?.addEventListener('click', () => {
  hideShapeMenu()
  deleteSelectedShape()
})
cropButton?.addEventListener('click', () => {
  page.cropping = !page.cropping
  render()
})
zoomInput?.addEventListener('input', () => {
  setZoom(Number(zoomInput.value) / 100)
})
zoomOut?.addEventListener('click', () => {
  setZoom((Math.round(page.zoom * 100) - 1) / 100)
})
zoomIn?.addEventListener('click', () => {
  setZoom((Math.round(page.zoom * 100) + 1) / 100)
})
if (paneSplit !== null && thumbs !== null) {
  const bar = paneSplit
  const preview = thumbs
  bar.addEventListener('pointerdown', (event) => {
    if (event.button !== 0) return
    event.preventDefault()
    const startX = event.clientX
    const startWidth = preview.getBoundingClientRect().width
    const pointerId = event.pointerId
    bar.classList.add('dragging')
    try {
      bar.setPointerCapture(pointerId)
    } catch {
      // A pointer the browser will not capture still reports move and up on window.
    }
    const move = (ev: PointerEvent): void => {
      if (ev.pointerId !== pointerId) return
      preview.style.width = `${thumbWidth(startWidth + ev.clientX - startX)}px`
    }
    const end = (ev: PointerEvent): void => {
      if (ev.pointerId !== pointerId) return
      bar.classList.remove('dragging')
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', end)
      window.removeEventListener('pointercancel', end)
    }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', end)
    window.addEventListener('pointercancel', end)
  })
}
window.addEventListener('resize', () => {
  if (thumbs === null) return
  const fitted = thumbWidth(thumbs.getBoundingClientRect().width)
  if (Math.abs(thumbs.getBoundingClientRect().width - fitted) > 1) thumbs.style.width = `${fitted}px`
})
// Wheel zooms only after the slide page itself is selected. A picture, a text box, or a click beside the page leaves the wheel for scrolling. Empty autoshapes count as the page.
viewport?.addEventListener(
  'wheel',
  (event) => {
    if (page.doc === null || !page.pageSelected) return
    event.preventDefault()
    setZoom(page.zoom * (event.deltaY < 0 ? 1.08 : 1 / 1.08))
  },
  { passive: false }
)
viewport?.addEventListener('pointerdown', (event) => {
  if (event.button !== 0 || event.target !== viewport) return
  if (!page.pageSelected && page.selectedId === undefined && !page.editing) return
  page.pageSelected = false
  page.selectedId = undefined
  page.editing = false
  page.cropping = false
  page.live = undefined
  render()
})
thumbs?.addEventListener('scroll', () => {
  hideMenu()
})
sheet?.addEventListener('pointerdown', (event) => {
  if (event.button !== 0 || canvas === null || !(event.target instanceof Node)) return
  if (event.target instanceof Element && event.target.closest('#handles, #shape-menu')) return
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
  focusSheet()
  const previous = page.selectedId
  const content = shape !== undefined && isContent(shape)
  page.pageSelected = !content
  page.selectedId = content ? shape.id : undefined
  page.editing = false
  page.cropping = content && shape.fill.kind === 'picture' && shape.id === previous && page.cropping
  if (content) {
    const shown = shownShape(shape)
    shapeGesture = {
      kind: 'pending',
      id: shape.id,
      pointerId: event.pointerId,
      startClientX: event.clientX,
      startClientY: event.clientY,
      origin: shapeBox(shown),
      handle: '',
      crop: pictureCrop(shown),
      moved: false
    }
    canvas.setPointerCapture(event.pointerId)
  }
  render()
})
sheet?.addEventListener('dblclick', (event) => {
  if (canvas === null || !(event.target instanceof Node)) return
  if (editor !== null && editor.contains(event.target)) return
  const box = canvas.getBoundingClientRect()
  const shape = hitTest(event.clientX - box.left, event.clientY - box.top)
  if (shape?.editable !== true) return
  page.pageSelected = false
  page.selectedId = shape.id
  page.editing = true
  page.cropping = false
  render()
  placeCaret(event.clientX, event.clientY)
})
sheet?.addEventListener('contextmenu', (event) => {
  if (canvas === null) return
  const box = canvas.getBoundingClientRect()
  const shape = hitTest(event.clientX - box.left, event.clientY - box.top)
  if (shape === undefined || !canMove(shape)) return
  if (!isContent(shape)) {
    event.preventDefault()
    page.pageSelected = true
    page.selectedId = undefined
    page.editing = false
    render()
    return
  }
  event.preventDefault()
  page.pageSelected = false
  page.selectedId = shape.id
  page.editing = false
  render()
  openShapeMenu(shape, event.clientX, event.clientY)
})
handles?.addEventListener('pointerdown', (event) => {
  if (event.button !== 0 || canvas === null) return
  const handle =
    event.target instanceof Element ? event.target.closest<HTMLElement>('[data-handle]') : null
  if (handle === null) return
  const shape = selected()
  if (shape === undefined || !hasFrameHandles(shape)) return
  event.preventDefault()
  event.stopPropagation()
  const shown = shownShape(shape)
  shapeGesture = {
    kind: page.cropping && shape.fill.kind === 'picture' ? 'crop' : 'resize',
    id: shape.id,
    pointerId: event.pointerId,
    startClientX: event.clientX,
    startClientY: event.clientY,
    origin: shapeBox(shown),
    handle: handle.dataset.handle ?? 'se',
    crop: pictureCrop(shown),
    moved: true
  }
  handle.setPointerCapture(event.pointerId)
})
sheet?.addEventListener('mousedown', (event) => {
  if (event.button !== 0 || !(event.target instanceof Node)) return
  if (editor !== null && (event.target === editor || editor.contains(event.target))) return
  event.preventDefault()
})
window.addEventListener('pointermove', (event) => {
  moveShapeGesture(event)
})
window.addEventListener('pointerup', (event) => {
  finishShapeGesture(event)
  if (!pressing) return
  pressing = false
  if (page.editing && editor !== null && document.activeElement !== editor) editor.focus()
})
window.addEventListener('pointercancel', () => {
  pressing = false
})
canvas?.addEventListener('mousemove', (event) => {
  const shape = hitTest(event.offsetX, event.offsetY)
  const selectedShape = selected()
  if (
    shape !== undefined &&
    selectedShape !== undefined &&
    shape.id === selectedShape.id &&
    canMove(shape) &&
    !page.editing
  ) {
    canvas.style.cursor = 'move'
    return
  }
  canvas.style.cursor =
    shape !== undefined &&
    canMove(shape) &&
    !page.editing &&
    (shape.chart !== undefined || shape.id === selectedShape?.id)
      ? 'move'
      : shape?.editable === true
        ? 'text'
        : 'default'
})
editor?.addEventListener('input', (event) => {
  if (event.isComposing) return
  commitEditor()
})
editor?.addEventListener('compositionend', () => {
  commitEditor()
})
editor?.addEventListener('paste', (event) => {
  event.preventDefault()
  insertPlain(event.clipboardData?.getData('text/plain') ?? '')
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
  void openDropped(event)
})

const requested = new URLSearchParams(location.search).get('open')
if (requested !== null && safeFileName(requested)) void openRequested(requested)

if (embeddedFrame()) {
  window.addEventListener('message', (event) => {
    if (event.source !== window.parent || event.origin !== location.origin) return
    const message = event.data as { type?: unknown; name?: unknown; bytes?: unknown }
    if (message?.type !== 'dsh-office-open' || !(message.bytes instanceof ArrayBuffer)) return
    const name = typeof message.name === 'string' && message.name.length > 0 ? message.name : 'presentation.pptx'
    void openBytes(new Uint8Array(message.bytes), name)
  })
  window.parent.postMessage({ type: 'dsh-office-ready' }, location.origin)
}

async function openRequested(name: string): Promise<void> {
  try {
    const response = await fetch(encodeURIComponent(name))
    if (!response.ok) throw new Error('pptx: 找不到这个文件')
    page.handle = null
    page.servedName = name
    await openBytes(new Uint8Array(await response.arrayBuffer()), name)
  } catch (error) {
    note(error instanceof Error ? error.message : '无法打开这个文件')
  }
}

function embeddedFrame(): boolean {
  try {
    return window.self !== window.top
  } catch {
    return true
  }
}

function filePicker():
  | ((options: {
      types?: Array<{ description?: string; accept: Record<string, string[]> }>
    }) => Promise<SavedFile[]>)
  | undefined {
  return (
    window as Window & {
      showOpenFilePicker?: (options: {
        types?: Array<{ description?: string; accept: Record<string, string[]> }>
      }) => Promise<SavedFile[]>
    }
  ).showOpenFilePicker
}

function readPlainOpen(): boolean {
  try {
    return sessionStorage.getItem('pptx-plain-open') === '1'
  } catch {
    return false
  }
}

function usePlainOpen(): void {
  plainOpen = true
  try {
    sessionStorage.setItem('pptx-plain-open', '1')
  } catch {
    // Private mode can reject storage. This page load still keeps the plain picker.
  }
}

function choosePlainFile(): void {
  if (fileInput === null) return
  fileInput.value = ''
  fileInput.click()
}

function accessDenied(error: unknown): boolean {
  if (
    error instanceof DOMException &&
    (error.name === 'NotAllowedError' || error.name === 'SecurityError')
  ) {
    return true
  }
  return error instanceof Error && error.message.includes('not allowed')
}

async function pickFile(): Promise<void> {
  const picker = filePicker()
  if (picker === undefined) {
    usePlainOpen()
    choosePlainFile()
    return
  }
  try {
    const [handle] = await picker({
      types: [
        {
          description: 'PowerPoint',
          accept: {
            'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx']
          }
        }
      ]
    })
    if (handle === undefined) return
    let file: File
    try {
      file = await handle.getFile()
    } catch (error) {
      if (!accessDenied(error)) throw error
      // This browser can show the picker but then refuses to read the handle.
      usePlainOpen()
      page.handle = null
      choosePlainFile()
      return
    }
    page.handle = handle
    page.servedName = null
    await openBytes(new Uint8Array(await file.arrayBuffer()), handle.name)
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') return
    if (accessDenied(error)) {
      usePlainOpen()
      choosePlainFile()
      return
    }
    note(error instanceof Error ? error.message : '无法打开这个文件')
  }
}

async function openDropped(event: DragEvent): Promise<void> {
  const item = event.dataTransfer?.items[0]
  const readable = item as
    | (DataTransferItem & { getAsFileSystemHandle?: () => Promise<SavedFile | null> })
    | undefined
  const handle = await readable?.getAsFileSystemHandle?.()
  const file = item?.getAsFile() ?? event.dataTransfer?.files[0]
  if (file === undefined) return
  page.handle = handle ?? null
  page.servedName = null
  await openBytes(new Uint8Array(await file.arrayBuffer()), file.name)
}

async function openFile(file: File): Promise<void> {
  page.handle = null
  page.servedName = null
  await openBytes(new Uint8Array(await file.arrayBuffer()), file.name)
}

async function insertLocalImage(file: File): Promise<void> {
  const doc = page.doc
  if (doc === null) return
  const extension = imageExtension(file.name)
  if (extension === undefined) {
    note('只支持 PNG、JPEG、GIF 和 WebP')
    return
  }
  let recorded = false
  try {
    const bytes = new Uint8Array(await file.arrayBuffer())
    const bitmap = await createImageBitmap(new Blob([toBuffer(bytes)]))
    const size = doc.slideSize()
    const frame = fittedPicture(size.width, size.height, bitmap.width, bitmap.height)
    rememberHistory()
    recorded = true
    const id = doc.insertPicture(page.slideIndex, { bytes, extension }, frame)
    const shape = doc.scene(page.slideIndex).shapes.find((item) => item.id === id)
    if (shape?.fill.kind === 'picture') page.images.set(shape.fill.imagePath, bitmap)
    else bitmap.close()
    page.selectedId = id
    page.pageSelected = false
    page.editing = false
    page.cropping = false
    markDirty()
    focusSheet()
    render()
  } catch (error) {
    if (recorded) history.pop()
    note(error instanceof Error ? error.message : '无法插入图片')
  }
}

function imageExtension(name: string): string | undefined {
  const extension = /\.([a-z0-9]+)$/i.exec(name)?.[1]?.toLowerCase()
  if (
    extension === 'png' ||
    extension === 'jpg' ||
    extension === 'jpeg' ||
    extension === 'gif' ||
    extension === 'webp'
  ) {
    return extension
  }
  return undefined
}

function fittedPicture(
  slideWidth: number,
  slideHeight: number,
  pixelWidth: number,
  pixelHeight: number
): { x: number; y: number; width: number; height: number } {
  const maxWidth = slideWidth * 0.45
  const maxHeight = slideHeight * 0.45
  let width = Math.max(1, pixelWidth) * EMU_PER_PIXEL
  let height = Math.max(1, pixelHeight) * EMU_PER_PIXEL
  const scale = Math.min(1, maxWidth / width, maxHeight / height)
  width = Math.max(1, Math.round(width * scale))
  height = Math.max(1, Math.round(height * scale))
  return {
    x: Math.max(0, Math.round((slideWidth - width) / 2)),
    y: Math.max(0, Math.round((slideHeight - height) / 2)),
    width,
    height
  }
}

async function openBytes(bytes: Uint8Array, label: string): Promise<void> {
  const saved = bytes.slice()
  try {
    page.doc = await openPptx(saved)
  } catch (error) {
    page.doc = null
    page.saved = null
    clearDirty()
    history.length = 0
    if (saveButton !== null) saveButton.disabled = true
    if (addButton !== null) addButton.disabled = true
    if (insertImageButton !== null) insertImageButton.disabled = true
    note(error instanceof Error ? error.message : '无法打开这个文件')
    render()
    return
  }
  page.saved = saved
  page.label = label
  page.slideIndex = 0
  page.selectedId = undefined
  page.pageSelected = false
  page.editing = false
  page.editorShape = undefined
  page.thumbIndex = -1
  page.cropping = false
  page.live = undefined
  page.clip = null
  setZoom(1, false)
  page.images.clear()
  history.length = 0
  clearDirty()
  if (saveButton !== null) saveButton.disabled = false
  if (addButton !== null) addButton.disabled = false
  if (insertImageButton !== null) insertImageButton.disabled = false
  note(label)
  await loadImages()
  render()
}

async function saveCurrent(): Promise<void> {
  const doc = page.doc
  if (doc === null || saveButton === null || saveButton.disabled) return
  saveButton.disabled = true
  try {
    const bytes = await doc.save()
    await writeCurrent(bytes)
    page.saved = bytes.slice()
    const index = page.slideIndex
    const selectedId = page.selectedId
    page.doc = await openPptx(bytes.slice())
    page.slideIndex = Math.min(index, Math.max(page.doc.slides().length - 1, 0))
    page.selectedId = selectedId
    page.editing = false
    page.editorShape = undefined
    history.length = 0
    clearDirty()
    note(page.label.length > 0 ? page.label : '已保存')
    await loadImages()
    render()
  } catch (error) {
    note(error instanceof Error ? error.message : '无法保存')
  } finally {
    if (page.doc !== null) saveButton.disabled = false
  }
}

async function writeCurrent(bytes: Uint8Array): Promise<void> {
  const body = new Blob([toBuffer(bytes)], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  })
  if (page.handle !== null) {
    const writable = await page.handle.createWritable()
    await writable.write(body)
    await writable.close()
    return
  }
  if (page.servedName !== null) {
    const response = await fetch(`/save?name=${encodeURIComponent(page.servedName)}`, {
      method: 'POST',
      body
    })
    if (!response.ok) throw new Error('无法覆盖当前文件')
    return
  }
  throw new Error('请用「打开」选择文件后再保存')
}

async function discardEdits(): Promise<void> {
  if (page.saved === null || !page.dirty) return
  const index = page.slideIndex
  page.doc = await openPptx(page.saved.slice())
  page.slideIndex = Math.min(index, Math.max(page.doc.slides().length - 1, 0))
  page.selectedId = undefined
  page.pageSelected = false
  page.editing = false
  page.editorShape = undefined
  page.thumbIndex = -1
  page.images.clear()
  history.length = 0
  clearDirty()
  note(page.label)
  await loadImages()
  render()
}

function addSlide(): void {
  const doc = page.doc
  if (doc === null) return
  rememberHistory()
  page.slideIndex = doc.addSlide(page.slideIndex)
  finishStructureChange()
}

function duplicateAt(index: number): void {
  const doc = page.doc
  if (doc === null || doc.slides()[index] === undefined) return
  rememberHistory()
  page.slideIndex = doc.duplicateSlide(index)
  finishStructureChange()
  void loadImages()
}

function deleteAt(index: number): void {
  const doc = page.doc
  if (doc === null || doc.slides().length <= 1 || doc.slides()[index] === undefined) return
  rememberHistory()
  doc.deleteSlide(index)
  page.slideIndex = Math.min(index, doc.slides().length - 1)
  finishStructureChange()
}

function finishStructureChange(): void {
  page.selectedId = undefined
  page.pageSelected = false
  page.editing = false
  page.editorShape = undefined
  page.thumbIndex = -1
  if (document.activeElement === editor) editor?.blur()
  markDirty()
  render()
}

function selectSlide(index: number): void {
  if (page.doc?.slides()[index] === undefined) return
  page.slideIndex = index
  page.selectedId = undefined
  page.pageSelected = false
  page.editing = false
  page.editorShape = undefined
  if (document.activeElement === editor) editor?.blur()
  render()
}

function openMenu(index: number, clientX: number, clientY: number): void {
  if (slideMenu === null || page.doc === null) return
  menuIndex = index
  if (deleteSlideButton !== null) deleteSlideButton.disabled = page.doc.slides().length <= 1
  slideMenu.hidden = false
  const width = slideMenu.offsetWidth
  const height = slideMenu.offsetHeight
  const left = Math.min(clientX, Math.max(8, window.innerWidth - width - 8))
  const top = Math.min(clientY, Math.max(8, window.innerHeight - height - 8))
  slideMenu.style.left = `${left}px`
  slideMenu.style.top = `${top}px`
}

function hideMenu(): void {
  if (slideMenu === null) return
  slideMenu.hidden = true
  menuIndex = -1
}

function markDirty(): void {
  page.dirty = true
  if (discardButton !== null) discardButton.disabled = false
}

function clearDirty(): void {
  page.dirty = false
  if (discardButton !== null) discardButton.disabled = true
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
  if (document.activeElement !== editor || switched) paintEditor(shape)
}

function commitEditor(): void {
  if (painting || editor === null) return
  const doc = page.doc
  const shape = selected()
  if (doc === null || shape === undefined || !shape.editable) return
  const grid = readEditorGrid(editor)
  const plain = editorPlain(editor)
  if (sameEditorText(shape, grid, plain)) return
  const offset = grid === undefined ? caretOffset(editor) : 0
  rememberHistory()
  try {
    if (grid === undefined) doc.setShapeText(page.slideIndex, shape.id, plain)
    else doc.setShapeRuns(page.slideIndex, shape.id, grid)
  } catch (error) {
    history.pop()
    note(error instanceof Error ? error.message : '无法修改文字')
    return
  }
  markDirty()
  render()
  if (grid === undefined) {
    const updated = selected()
    if (updated !== undefined) paintEditor(updated)
    restoreCaret(editor, offset)
  }
}

function placeEditor(shape: PaintShape): void {
  if (editorFrame === null || editor === null) return
  const scale = page.scale
  const text = shape.text
  const paragraph = text?.paragraphs[0]
  editorFrame.style.left = `${(shape.x / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.top = `${(shape.y / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.width = `${(shape.width / EMU_PER_PIXEL) * scale}px`
  const overflowing = text?.clip === false
  editorFrame.classList.toggle('overflowing', overflowing)
  editorFrame.style.overflow = overflowing ? 'visible' : 'hidden'
  editorFrame.style.height = overflowing ? 'auto' : `${(shape.height / EMU_PER_PIXEL) * scale}px`
  editorFrame.style.minHeight = overflowing ? `${(shape.height / EMU_PER_PIXEL) * scale}px` : ''
  editorFrame.style.transform = `rotate(${shape.rotation}deg)`
  const pad = text?.padding
  editorFrame.style.padding =
    pad === undefined
      ? '0'
      : `${(pad.top / EMU_PER_PIXEL) * scale}px ${(pad.right / EMU_PER_PIXEL) * scale}px ${(pad.bottom / EMU_PER_PIXEL) * scale}px ${(pad.left / EMU_PER_PIXEL) * scale}px`
  const anchor = text?.anchor ?? 't'
  editorFrame.style.justifyContent =
    anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start'
  editor.style.fontSize = ''
  editor.style.fontWeight = ''
  editor.style.fontStyle = ''
  editor.style.fontFamily = ''
  editor.style.color = ''
  editor.style.textDecoration = ''
  editor.style.textAlign =
    paragraph?.align === 'ctr' ? 'center' : paragraph?.align === 'r' ? 'right' : 'left'
  editor.style.lineHeight = 'normal'
}

function paintEditor(shape: PaintShape): void {
  const grid = (shape.text?.paragraphs ?? []).map((paragraph) =>
    paragraph.runs.map((run) => run.text)
  )
  paintGrid(shape, grid)
}

function paintGrid(shape: PaintShape, grid: readonly (readonly string[])[]): void {
  if (editor === null) return
  const paragraphs = shape.text?.paragraphs ?? []
  const scale = page.scale
  const nodes: HTMLDivElement[] = []
  painting = true
  paragraphs.forEach((paragraph, paragraphIndex) => {
    const block = document.createElement('div')
    block.dataset.p = String(paragraphIndex)
    const box = paragraph.runs.reduce((max, run) => Math.max(max, lineBox(run, paragraph)), 0)
    block.style.marginTop = `${((paragraph.spaceBefore * 96) / 72) * scale}px`
    block.style.marginBottom = `${((paragraph.spaceAfter * 96) / 72) * scale}px`
    block.style.lineHeight = `${(box || 1) * scale}px`
    block.style.textAlign =
      paragraph.align === 'ctr' ? 'center' : paragraph.align === 'r' ? 'right' : 'left'
    block.style.whiteSpace = 'pre-wrap'
    paragraph.runs.forEach((run, runIndex) => {
      const span = document.createElement('span')
      span.dataset.p = String(paragraphIndex)
      span.dataset.r = String(runIndex)
      styleRun(span, run)
      span.textContent = grid[paragraphIndex]?.[runIndex] ?? run.text
      block.append(span)
    })
    nodes.push(block)
  })
  editor.replaceChildren(...nodes)
  painting = false
}

function styleRun(span: HTMLSpanElement, run: PaintRun): void {
  const scale = page.scale
  span.style.fontSize = `${((run.fontSize * 96) / 72) * scale}px`
  span.style.fontWeight = run.bold ? '700' : '400'
  span.style.fontStyle = run.italic ? 'italic' : 'normal'
  span.style.fontFamily = editorFamily(run)
  span.style.color = run.color
  span.style.textDecoration = run.underline ? 'underline' : 'none'
  span.style.whiteSpace = 'pre-wrap'
}

function readEditorGrid(root: HTMLElement): string[][] | undefined {
  const grid: string[][] = []
  for (const node of root.childNodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      if ((node.textContent ?? '').length === 0) continue
      return undefined
    }
    if (!(node instanceof HTMLDivElement)) return undefined
    const paragraphIndex = Number(node.dataset.p)
    if (!Number.isInteger(paragraphIndex) || grid[paragraphIndex] !== undefined) return undefined
    const row: string[] = []
    for (const child of node.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        if ((child.textContent ?? '').length === 0) continue
        return undefined
      }
      if (!(child instanceof HTMLSpanElement)) return undefined
      const runIndex = Number(child.dataset.r)
      if (!Number.isInteger(runIndex) || row[runIndex] !== undefined) return undefined
      row[runIndex] = child.textContent ?? ''
    }
    if (row.length === 0) row.push(node.textContent ?? '')
    for (let run = 0; run < row.length; run += 1) {
      if (row[run] === undefined) return undefined
    }
    grid[paragraphIndex] = row
  }
  for (let index = 0; index < grid.length; index += 1) {
    if (grid[index] === undefined) return undefined
  }
  return grid
}

function editorPlain(root: HTMLElement): string {
  const blocks = [...root.children].filter(
    (node): node is HTMLElement => node instanceof HTMLElement
  )
  if (blocks.length === 0) return (root.textContent ?? '').replace(/\r\n/g, '\n')
  return blocks.map((block) => block.textContent ?? '').join('\n')
}

function sameEditorText(
  shape: PaintShape,
  grid: string[][] | undefined,
  plain: string
): boolean {
  if (grid === undefined) return plain === shape.plain
  const paragraphs = shape.text?.paragraphs ?? []
  if (grid.length !== paragraphs.length) return false
  return paragraphs.every((paragraph, index) => {
    const row = grid[index]
    return (
      row !== undefined &&
      row.length === paragraph.runs.length &&
      paragraph.runs.every((run, runIndex) => run.text === row[runIndex])
    )
  })
}

function caretOffset(root: HTMLElement): number {
  const selection = document.getSelection()
  if (selection === null || selection.rangeCount === 0 || !root.contains(selection.anchorNode))
    return 0
  const range = selection.getRangeAt(0).cloneRange()
  range.selectNodeContents(root)
  range.setEnd(selection.getRangeAt(0).startContainer, selection.getRangeAt(0).startOffset)
  return range.toString().length
}

function restoreCaret(root: HTMLElement, offset: number): void {
  const selection = document.getSelection()
  if (selection === null) return
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let remaining = offset
  let node = walker.nextNode()
  while (node !== null) {
    const length = node.textContent?.length ?? 0
    if (remaining <= length) {
      const range = document.createRange()
      range.setStart(node, remaining)
      range.collapse(true)
      selection.removeAllRanges()
      selection.addRange(range)
      return
    }
    remaining -= length
    node = walker.nextNode()
  }
  const range = document.createRange()
  range.selectNodeContents(root)
  range.collapse(false)
  selection.removeAllRanges()
  selection.addRange(range)
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

function typingTarget(): boolean {
  const node = document.activeElement
  return node === editor || node instanceof HTMLInputElement || node instanceof HTMLTextAreaElement
}

function fieldTarget(): boolean {
  const node = document.activeElement
  return node instanceof HTMLInputElement || node instanceof HTMLTextAreaElement
}

function focusSheet(): void {
  sheet?.focus({ preventScroll: true })
}

function rememberHistory(): void {
  if (page.doc === null || applyingHistory) return
  history.push(page.doc.capture())
  if (history.length > 50) history.shift()
}

function undo(): void {
  const doc = page.doc
  const checkpoint = history.pop()
  if (doc === null || checkpoint === undefined) return
  applyingHistory = true
  try {
    doc.restore(checkpoint)
    if (checkpoint.dirty) markDirty()
    else clearDirty()
    const current = doc.scene(page.slideIndex).shapes.find((shape) => shape.id === page.selectedId)
    if (current === undefined) {
      page.selectedId = undefined
      page.editing = false
    }
    page.cropping = false
    page.live = undefined
    page.editorShape = undefined
    page.images.clear()
  } finally {
    applyingHistory = false
  }
  void loadImages().then(() => {
    render()
  })
}

const THUMB_MIN = 96
const STAGE_MIN = 160

function thumbWidth(next: number): number {
  const total = thumbs?.parentElement?.getBoundingClientRect().width ?? 0
  const split = paneSplit?.getBoundingClientRect().width ?? 0
  const max = Math.max(THUMB_MIN, total - split - STAGE_MIN)
  return Math.round(Math.min(max, Math.max(THUMB_MIN, next)))
}

function setZoom(next: number, paint = true): void {
  page.zoom = Math.min(4, Math.max(0.1, next))
  const percent = Math.round(page.zoom * 100)
  if (zoomInput !== null) zoomInput.value = String(percent)
  if (zoomLabel !== null) zoomLabel.textContent = `${percent}%`
  if (zoomOut !== null) zoomOut.disabled = percent <= 10
  if (zoomIn !== null) zoomIn.disabled = percent >= 400
  if (paint && page.doc !== null) render()
}

function canMove(shape: PaintShape): boolean {
  return (
    shape.selectable &&
    (shape.fill.kind === 'picture' || shape.editable || shape.chart !== undefined)
  )
}

function isContent(shape: PaintShape): boolean {
  return (
    canMove(shape) &&
    (shape.fill.kind === 'picture' || shape.plain.length > 0 || shape.chart !== undefined)
  )
}

function hasFrameHandles(shape: PaintShape): boolean {
  return shape.fill.kind === 'picture' || shape.chart !== undefined
}

function canCopy(shape: PaintShape): boolean {
  return shape.fill.kind === 'picture' || shape.chart !== undefined || shape.editable
}

function shapeBox(shape: PaintShape): Box {
  return {
    x: shape.x / EMU_PER_PIXEL,
    y: shape.y / EMU_PER_PIXEL,
    width: shape.width / EMU_PER_PIXEL,
    height: shape.height / EMU_PER_PIXEL
  }
}

function pictureCrop(shape: PaintShape): {
  left: number
  top: number
  right: number
  bottom: number
} {
  if (shape.fill.kind !== 'picture' || shape.fill.crop === undefined) {
    return { left: 0, top: 0, right: 0, bottom: 0 }
  }
  return { ...shape.fill.crop }
}

function shownShape(shape: PaintShape): PaintShape {
  const live = page.live
  if (live === undefined || live.id !== shape.id) return shape
  const framed = {
    ...shape,
    x: live.box.x * EMU_PER_PIXEL,
    y: live.box.y * EMU_PER_PIXEL,
    width: live.box.width * EMU_PER_PIXEL,
    height: live.box.height * EMU_PER_PIXEL
  }
  if (live.crop === undefined || shape.fill.kind !== 'picture') return framed
  return { ...framed, fill: { kind: 'picture', imagePath: shape.fill.imagePath, crop: live.crop } }
}

function moveShapeGesture(event: PointerEvent): void {
  const gesture = shapeGesture
  if (gesture === null || event.pointerId !== gesture.pointerId) return
  const dx = (event.clientX - gesture.startClientX) / page.zoom
  const dy = (event.clientY - gesture.startClientY) / page.zoom
  if (gesture.kind === 'pending') {
    if (Math.hypot(dx, dy) < 4) return
    gesture.kind = 'move'
    gesture.moved = true
    page.editing = false
  }
  if (gesture.kind === 'move') {
    page.live = {
      id: gesture.id,
      box: {
        x: gesture.origin.x + dx,
        y: gesture.origin.y + dy,
        width: gesture.origin.width,
        height: gesture.origin.height
      }
    }
  } else if (gesture.kind === 'resize') {
    page.live = { id: gesture.id, box: resizeBox(gesture.origin, gesture.handle, dx, dy) }
  } else if (gesture.kind === 'crop') {
    page.live = {
      id: gesture.id,
      box: gesture.origin,
      crop: cropFromPointer(
        gesture.origin,
        gesture.crop,
        gesture.handle,
        event.clientX,
        event.clientY
      )
    }
  }
  render()
}

function finishShapeGesture(event: PointerEvent): void {
  const gesture = shapeGesture
  if (gesture === null || event.pointerId !== gesture.pointerId) return
  shapeGesture = null
  const live = page.live
  page.live = undefined
  const doc = page.doc
  if (doc === null || gesture.kind === 'pending' || live === undefined) {
    render()
    return
  }
  rememberHistory()
  if (gesture.kind === 'crop' && live.crop !== undefined) {
    doc.setPictureCrop(page.slideIndex, gesture.id, live.crop)
  } else {
    doc.setShapeFrame(page.slideIndex, gesture.id, {
      x: live.box.x * EMU_PER_PIXEL,
      y: live.box.y * EMU_PER_PIXEL,
      width: live.box.width * EMU_PER_PIXEL,
      height: live.box.height * EMU_PER_PIXEL
    })
  }
  markDirty()
  render()
}

function resizeBox(origin: Box, handle: string, dx: number, dy: number): Box {
  const min = 8
  let left = origin.x
  let right = origin.x + origin.width
  let top = origin.y
  let bottom = origin.y + origin.height
  if (handle.includes('w')) left = Math.min(origin.x + dx, right - min)
  if (handle.includes('e')) right = Math.max(origin.x + origin.width + dx, left + min)
  if (handle.includes('n')) top = Math.min(origin.y + dy, bottom - min)
  if (handle.includes('s')) bottom = Math.max(origin.y + origin.height + dy, top + min)
  return {
    x: left,
    y: top,
    width: Math.max(min, right - left),
    height: Math.max(min, bottom - top)
  }
}

function cropFromPointer(
  origin: Box,
  start: { left: number; top: number; right: number; bottom: number },
  handle: string,
  clientX: number,
  clientY: number
): { left: number; top: number; right: number; bottom: number } {
  const point = slidePoint(clientX, clientY)
  const crop = { ...start }
  const remain = 0.08
  if (handle.includes('w')) {
    crop.left = clamp((point.x - origin.x) / origin.width, 0, 1 - start.right - remain)
  }
  if (handle.includes('e')) {
    crop.right = clamp(
      (origin.x + origin.width - point.x) / origin.width,
      0,
      1 - start.left - remain
    )
  }
  if (handle.includes('n')) {
    crop.top = clamp((point.y - origin.y) / origin.height, 0, 1 - start.bottom - remain)
  }
  if (handle.includes('s')) {
    crop.bottom = clamp(
      (origin.y + origin.height - point.y) / origin.height,
      0,
      1 - start.top - remain
    )
  }
  return crop
}

function slidePoint(clientX: number, clientY: number): { x: number; y: number } {
  const rect = canvas?.getBoundingClientRect()
  return {
    x: (clientX - (rect?.left ?? 0)) / page.zoom,
    y: (clientY - (rect?.top ?? 0)) / page.zoom
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function syncHandles(shape: PaintShape | undefined): void {
  if (handles === null) return
  if (shape === undefined || !hasFrameHandles(shape) || page.editing) {
    handles.hidden = true
    return
  }
  const box = shapeBox(shape)
  handles.hidden = false
  const picture = shape.fill.kind === 'picture'
  handles.classList.toggle('cropping', page.cropping && picture)
  if (cropButton !== null) cropButton.hidden = !picture
  handles.style.left = `${box.x * page.zoom}px`
  handles.style.top = `${box.y * page.zoom}px`
  handles.style.width = `${box.width * page.zoom}px`
  handles.style.height = `${box.height * page.zoom}px`
  cropButton?.setAttribute('aria-pressed', page.cropping ? 'true' : 'false')
}

function openShapeMenu(shape: PaintShape, clientX: number, clientY: number): void {
  if (shapeMenu === null) return
  if (copyShapeButton !== null) copyShapeButton.hidden = !canCopy(shape)
  shapeMenu.hidden = false
  const width = shapeMenu.offsetWidth
  const height = shapeMenu.offsetHeight
  shapeMenu.style.left = `${Math.min(clientX, Math.max(8, window.innerWidth - width - 8))}px`
  shapeMenu.style.top = `${Math.min(clientY, Math.max(8, window.innerHeight - height - 8))}px`
}

function hideShapeMenu(): void {
  if (shapeMenu === null) return
  shapeMenu.hidden = true
}

function deleteSelectedShape(): void {
  const doc = page.doc
  const shape = selected()
  if (doc === null || shape === undefined || !canMove(shape)) return
  rememberHistory()
  doc.deleteShape(page.slideIndex, shape.id)
  page.selectedId = undefined
  page.editing = false
  page.cropping = false
  page.live = undefined
  markDirty()
  render()
}

function copySelectedPicture(): boolean {
  const shape = selected()
  if (shape === undefined || !canCopy(shape)) return false
  page.clip = { slide: page.slideIndex, id: shape.id }
  page.pasteCount = 0
  const bitmap = shape.fill.kind === 'picture' ? page.images.get(shape.fill.imagePath) : undefined
  if (bitmap !== undefined && navigator.clipboard !== undefined) {
    const board = document.createElement('canvas')
    board.width = bitmap.width
    board.height = bitmap.height
    board.getContext('2d')?.drawImage(bitmap, 0, 0)
    void board.toBlob((blob) => {
      if (blob === null || navigator.clipboard.write === undefined) return
      void navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]).catch(() => {
        // The in-app paste still uses the captured shape when the browser refuses the clipboard.
      })
    }, 'image/png')
  }
  return true
}

function pastePicture(): void {
  const doc = page.doc
  const clip = page.clip
  if (doc === null || clip === null) return
  page.pasteCount += 1
  const offset = page.pasteCount * 12 * EMU_PER_PIXEL
  rememberHistory()
  try {
    page.selectedId = doc.duplicateShape(clip.slide, clip.id, page.slideIndex, offset, offset)
  } catch (error) {
    history.pop()
    page.pasteCount -= 1
    note(error instanceof Error ? error.message : '无法粘贴')
    return
  }
  page.pageSelected = false
  page.editing = false
  page.cropping = false
  markDirty()
  render()
}

const textInk = new Map<string, { top: number; bottom: number }>()

function rememberTextInk(scene: SlideScene): void {
  textInk.clear()
  const ctx = context
  if (ctx === null || ctx === undefined) return
  for (const shape of scene.shapes) {
    const text = shape.text
    if (text === undefined || text.clip) continue
    const width = shape.width / EMU_PER_PIXEL
    const height = shape.height / EMU_PER_PIXEL
    const y = shape.y / EMU_PER_PIXEL
    const padTop = text.padding.top / EMU_PER_PIXEL
    const padBottom = text.padding.bottom / EMU_PER_PIXEL
    const innerWidth = Math.max(
      1,
      width - text.padding.left / EMU_PER_PIXEL - text.padding.right / EMU_PER_PIXEL
    )
    const innerHeight = Math.max(0, height - padTop - padBottom)
    const contentHeight = text.paragraphs.reduce(
      (sum, paragraph) => sum + layoutParagraph(ctx, paragraph, innerWidth).height,
      0
    )
    let top = y + padTop
    if (text.anchor === 'ctr') top += Math.max(0, (innerHeight - contentHeight) / 2)
    if (text.anchor === 'b') top += Math.max(0, innerHeight - contentHeight)
    textInk.set(shape.id, { top, bottom: top + contentHeight })
  }
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
  const ink = textInk.get(shape.id)
  const top = ink === undefined ? y : Math.min(y, ink.top)
  const bottom = ink === undefined ? y + height : Math.max(y + height, ink.bottom)
  const centerY = y + height / 2
  const dx = px - (x + width / 2)
  const dy = py - centerY
  const rad = (shape.rotation * Math.PI) / 180
  const cos = Math.cos(rad)
  const sin = Math.sin(rad)
  const localX = dx * cos + dy * sin
  const localY = -dx * sin + dy * cos
  return Math.abs(localX) <= width / 2 && localY >= top - centerY && localY <= bottom - centerY
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
    for (const button of thumbButtons()) button.remove()
    if (addButton !== null) addButton.disabled = true
    if (handles !== null) handles.hidden = true
    sheet?.classList.remove('page-selected')
    syncEditor(undefined)
    return
  }
  syncThumbs(doc)
  sheet?.classList.toggle('page-selected', page.pageSelected)
  const scene = doc.scene(page.slideIndex)
  const width = Math.max(1, scene.width / EMU_PER_PIXEL)
  const height = Math.max(1, scene.height / EMU_PER_PIXEL)
  const scale = page.zoom
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
  const shown = shape === undefined ? undefined : shownShape(shape)
  rememberTextInk(scene)
  if (shown !== undefined && shown.selectable) strokeSelection(context, shown)
  syncHandles(shown)
  syncEditor(page.editing ? shape : undefined)
}

function syncThumbs(doc: PptxDocument): void {
  if (thumbs === null) return
  const slides = doc.slides()
  const buttons = thumbButtons()
  while (buttons.length > slides.length) buttons.pop()?.remove()
  const slideChanged = page.thumbIndex !== page.slideIndex
  slides.forEach((_, index) => {
    let button = buttons[index]
    if (button === undefined) {
      button = createThumb()
      thumbs.append(button)
      buttons.push(button)
    }
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

function thumbButtons(): HTMLButtonElement[] {
  if (thumbs === null) return []
  return [...thumbs.querySelectorAll<HTMLButtonElement>('button.thumb')]
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
    const index = Number(button.dataset.index)
    if (!Number.isInteger(index)) return
    drag = { index, pointerId: event.pointerId, startY: event.clientY, moved: false }
    button.setPointerCapture(event.pointerId)
  })
  button.addEventListener('pointermove', (event: PointerEvent) => {
    if (drag === null || event.pointerId !== drag.pointerId) return
    if (!drag.moved && Math.abs(event.clientY - drag.startY) < 4) return
    drag.moved = true
    button.classList.add('dragging')
    showDrop(dropSlot(event.clientY, drag.index))
    moveThumbGhost(event.clientX, event.clientY)
  })
  button.addEventListener('pointerup', (event: PointerEvent) => {
    if (drag === null || event.pointerId !== drag.pointerId) return
    finishDrag(event.clientY)
  })
  button.addEventListener('pointercancel', (event: PointerEvent) => {
    if (drag === null || event.pointerId !== drag.pointerId) return
    drag = null
    clearDrop()
  })
  button.addEventListener('contextmenu', (event: MouseEvent) => {
    event.preventDefault()
    const index = Number(button.dataset.index)
    if (!Number.isInteger(index)) return
    selectSlide(index)
    openMenu(index, event.clientX, event.clientY)
  })
  return button
}

function finishDrag(clientY: number): void {
  const state = drag
  drag = null
  clearDrop()
  if (state === null) return
  if (!state.moved) {
    selectSlide(state.index)
    return
  }
  const to = dropSlot(clientY, state.index).to
  const doc = page.doc
  if (doc === null || to === state.index) {
    selectSlide(state.index)
    return
  }
  doc.moveSlide(state.index, to)
  page.slideIndex = to
  finishStructureChange()
}

function dropSlot(clientY: number, from: number): { raw: number; to: number } {
  const buttons = thumbButtons()
  let raw = buttons.length
  for (let index = 0; index < buttons.length; index += 1) {
    const rect = buttons[index]?.getBoundingClientRect()
    if (rect !== undefined && clientY < rect.top + rect.height / 2) {
      raw = index
      break
    }
  }
  return { raw, to: from < raw ? raw - 1 : raw }
}

function showDrop(slot: { raw: number; to: number }): void {
  clearDrop()
  const buttons = thumbButtons()
  const from = drag?.index
  if (from === undefined) return
  const source = buttons[from]
  source?.classList.add('dragging')
  const height = (source?.offsetHeight ?? 84) + 14
  if (slot.to !== from) {
    buttons.forEach((button, index) => {
      let shift = 0
      if (index !== from && from < slot.raw && index > from && index < slot.raw) shift = -height
      if (index !== from && slot.raw < from && index >= slot.raw && index < from) shift = height
      button.style.transform = shift === 0 ? '' : `translateY(${shift}px)`
    })
    if (slot.raw >= buttons.length) buttons[buttons.length - 1]?.classList.add('drop-after')
    else buttons[slot.raw]?.classList.add('drop-before')
  }
  placeThumbGhost()
}

function placeThumbGhost(): void {
  if (thumbGhost === null || drag === null) return
  const source = thumbButtons()[drag.index]?.querySelector('canvas')
  if (!(source instanceof HTMLCanvasElement)) return
  if (thumbGhost.dataset.index !== String(drag.index)) {
    const copy = document.createElement('canvas')
    copy.width = source.width
    copy.height = source.height
    copy.style.width = source.style.width
    copy.style.height = source.style.height
    copy.getContext('2d')?.drawImage(source, 0, 0)
    thumbGhost.replaceChildren(copy)
    thumbGhost.dataset.index = String(drag.index)
  }
  thumbGhost.hidden = false
}

function moveThumbGhost(clientX: number, clientY: number): void {
  if (thumbGhost === null) return
  const width = thumbGhost.offsetWidth || 148
  thumbGhost.style.left = `${clientX - width / 2}px`
  thumbGhost.style.top = `${clientY - 24}px`
}

function clearDrop(): void {
  if (thumbs === null) return
  for (const button of thumbButtons()) {
    button.classList.remove('dragging', 'drop-before', 'drop-after')
    button.style.transform = ''
  }
  if (thumbGhost !== null) {
    thumbGhost.hidden = true
    delete thumbGhost.dataset.index
  }
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
  const drawn = shownShape(shape)
  const x = drawn.x / EMU_PER_PIXEL
  const y = drawn.y / EMU_PER_PIXEL
  const width = drawn.width / EMU_PER_PIXEL
  const height = drawn.height / EMU_PER_PIXEL
  ctx.save()
  const cx = x + width / 2
  const cy = y + height / 2
  ctx.translate(cx, cy)
  ctx.rotate((drawn.rotation * Math.PI) / 180)
  if (drawn.flipH || drawn.flipV) ctx.scale(drawn.flipH ? -1 : 1, drawn.flipV ? -1 : 1)
  ctx.translate(-cx, -cy)
  if (drawn.geometry === 'line') {
    strokeLine(ctx, drawn, x, y, width, height)
    ctx.restore()
    return
  }
  if (drawn.chart !== undefined) {
    trace(ctx, drawn, x, y, width, height)
    if (drawn.fill.kind === 'solid' || drawn.fill.kind === 'linear' || drawn.fill.kind === 'radial') {
      ctx.fillStyle = paintFillStyle(ctx, drawn.fill, x, y, width, height)
      ctx.fill()
    }
    drawChart(ctx, drawn, x, y, width, height)
    if (drawn.line !== undefined) {
      ctx.strokeStyle = drawn.line.color
      ctx.lineWidth = Math.max(drawn.line.width / EMU_PER_PIXEL, 0.5)
      setDash(ctx, drawn.line.dash)
      trace(ctx, drawn, x, y, width, height)
      ctx.stroke()
      ctx.setLineDash([])
    }
    ctx.restore()
    return
  }
  if (drawn.shadow !== undefined) {
    ctx.shadowColor = drawn.shadow.color
    ctx.shadowBlur = drawn.shadow.blur / EMU_PER_PIXEL
    ctx.shadowOffsetX = drawn.shadow.offsetX / EMU_PER_PIXEL
    ctx.shadowOffsetY = drawn.shadow.offsetY / EMU_PER_PIXEL
  }
  trace(ctx, drawn, x, y, width, height)
  if (drawn.fill.kind === 'solid' || drawn.fill.kind === 'linear' || drawn.fill.kind === 'radial') {
    ctx.fillStyle = paintFillStyle(ctx, drawn.fill, x, y, width, height)
    ctx.fill()
  } else if (drawn.shadow !== undefined && drawn.fill.kind === 'picture') {
    ctx.fillStyle = '#ffffff'
    ctx.fill()
  }
  ctx.shadowColor = 'transparent'
  if (drawn.fill.kind === 'picture') {
    ctx.save()
    trace(ctx, drawn, x, y, width, height)
    ctx.clip()
    const bitmap = page.images.get(drawn.fill.imagePath)
    if (bitmap === undefined) {
      ctx.fillStyle = '#e6e8ea'
      ctx.fillRect(x, y, width, height)
    } else {
      drawBitmap(ctx, bitmap, drawn.fill.crop, x, y, width, height)
    }
    ctx.restore()
  }
  if (drawn.line !== undefined) {
    ctx.strokeStyle = drawn.line.color
    ctx.lineWidth = Math.max(drawn.line.width / EMU_PER_PIXEL, 0.5)
    setDash(ctx, drawn.line.dash)
    trace(ctx, drawn, x, y, width, height)
    ctx.stroke()
    ctx.setLineDash([])
  }
  if (drawn.text !== undefined && drawn.id !== omitTextId) drawText(ctx, drawn, x, y, width, height)
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

function drawChart(
  ctx: CanvasRenderingContext2D,
  shape: PaintShape,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  const chart = shape.chart
  if (chart === undefined) return
  const titleSize = 16
  const titleHeight = chart.title.length > 0 ? titleSize + 10 : 0
  ctx.fillStyle = '#12283d'
  ctx.font = `700 ${titleSize}px "Microsoft YaHei", "Segoe UI", sans-serif`
  ctx.textAlign = 'center'
  ctx.textBaseline = 'top'
  if (chart.title.length > 0) ctx.fillText(chart.title, x + width / 2, y + 2)
  const legend = chart.legend ? layoutLegend(ctx, chart.points, width) : { height: 0, rows: [] }
  const plotTop = y + titleHeight
  const plotHeight = Math.max(1, height - titleHeight - legend.height - 8)
  const outer = Math.max(1, Math.min(width, plotHeight) / 2 - 8)
  const inner = outer * chart.hole
  const cx = x + width / 2
  const cy = plotTop + plotHeight / 2
  const total = chart.points.reduce((sum, point) => sum + point.value, 0)
  if (total <= 0) return
  let angle = -Math.PI / 2 + (chart.start * Math.PI) / 180
  for (const point of chart.points) {
    const sweep = (point.value / total) * Math.PI * 2
    ctx.beginPath()
    if (inner <= 1) {
      ctx.moveTo(cx, cy)
      ctx.arc(cx, cy, outer, angle, angle + sweep)
    } else {
      ctx.arc(cx, cy, outer, angle, angle + sweep)
      ctx.arc(cx, cy, inner, angle + sweep, angle, true)
    }
    ctx.closePath()
    ctx.fillStyle = point.color
    ctx.fill()
    ctx.strokeStyle = '#ffffff'
    ctx.lineWidth = 2
    ctx.stroke()
    angle += sweep
  }
  if (legend.height > 0) drawLegend(ctx, legend, chart.points, x, y + height - legend.height, width)
  ctx.textAlign = 'left'
}

interface LegendRow {
  readonly items: readonly number[]
  readonly width: number
}

function layoutLegend(
  ctx: CanvasRenderingContext2D,
  points: NonNullable<PaintShape['chart']>['points'],
  width: number
): { height: number; rows: LegendRow[] } {
  ctx.font = '13px "Microsoft YaHei", "Segoe UI", sans-serif'
  const rows: { items: number[]; width: number }[] = []
  let row: { items: number[]; width: number } = { items: [], width: 0 }
  points.forEach((point, index) => {
    const itemWidth = 14 + ctx.measureText(point.label).width + 16
    if (row.items.length > 0 && row.width + itemWidth > width) {
      rows.push(row)
      row = { items: [], width: 0 }
    }
    row.items.push(index)
    row.width += itemWidth
  })
  if (row.items.length > 0) rows.push(row)
  return { height: rows.length * 20 + 8, rows }
}

function drawLegend(
  ctx: CanvasRenderingContext2D,
  legend: { readonly rows: readonly LegendRow[] },
  points: NonNullable<PaintShape['chart']>['points'],
  x: number,
  y: number,
  width: number
): void {
  ctx.font = '13px "Microsoft YaHei", "Segoe UI", sans-serif'
  ctx.textAlign = 'left'
  ctx.textBaseline = 'middle'
  ctx.fillStyle = '#5b6b7c'
  legend.rows.forEach((row, rowIndex) => {
    let cursor = x + Math.max(0, (width - row.width) / 2)
    const top = y + rowIndex * 20 + 10
    for (const index of row.items) {
      const point = points[index]
      if (point === undefined) continue
      ctx.fillStyle = point.color
      ctx.fillRect(cursor, top - 5, 10, 10)
      ctx.fillStyle = '#5b6b7c'
      ctx.fillText(point.label, cursor + 14, top)
      cursor += 14 + ctx.measureText(point.label).width + 16
    }
  })
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
  if (text.clip) {
    trace(ctx, shape, x, y, width, height)
    ctx.clip()
  }
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
    const baseline = sharedBaseline(ctx, paragraph, line, cursor, index)
    if (index === 0 && paragraph.bullet !== undefined && paragraph.bulletRun !== undefined) {
      ctx.font = fontOf(paragraph.bulletRun, isCjk(paragraph.bullet))
      ctx.fillStyle = paragraph.bulletRun.color
      ctx.fillText(paragraph.bullet, textX + paragraph.indent, baseline)
    }
    for (const token of line.tokens) {
      ctx.font = fontOf(token.run, token.cjk)
      ctx.fillStyle = token.run.color
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

function sharedBaseline(
  ctx: CanvasRenderingContext2D,
  paragraph: LaidParagraph,
  line: LaidLine,
  cursor: number,
  index: number
): number {
  let ascent = 0
  let descent = 0
  for (const token of line.tokens) {
    ascent = Math.max(ascent, token.ascent)
    descent = Math.max(descent, token.descent)
  }
  if (index === 0 && paragraph.bullet !== undefined && paragraph.bulletRun !== undefined) {
    ctx.font = fontOf(paragraph.bulletRun, isCjk(paragraph.bullet))
    const size = (paragraph.bulletRun.fontSize * 96) / 72
    const edge = fontEdges(ctx.measureText(paragraph.bullet), size)
    ascent = Math.max(ascent, edge.ascent)
    descent = Math.max(descent, edge.descent)
  }
  if (ascent === 0 && descent === 0) return cursor + line.height * 0.8
  return lineBaseline(cursor, line.height, ascent, descent)
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

function toBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}
