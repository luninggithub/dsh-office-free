import type { Univer } from '@univerjs/core'
import {
  ICommandService,
  IUniverInstanceService,
  LifecycleService,
  LifecycleStages,
  UniverInstanceType
} from '@univerjs/core'
import { getNextWheelZoomRatio } from '@univerjs/engine-render'
import type { SlideModel } from '@univerjs-pro/slides'
import { ISlideDrawingService, SetSlideZoomRatioOperation } from '@univerjs-pro/slides'
import type { ISlideDrawingStateSnapshot } from '@univerjs-pro/slides-ui'
import { ISlideDrawingStateService } from '@univerjs-pro/slides-ui'
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { filter, take } from 'rxjs'
import { currentLang } from '../i18n/index.ts'
import { readLocation } from './config.ts'
import {
  canAddRegion,
  canApply,
  createDefaultRegion,
  createDragRegion,
  createPointRegion,
  isDirty,
  isDrag,
  moveRegion,
  nextRegionId,
  promptOf,
  regionFrame,
  resizeRegion,
  type Handle,
  type ImageRegion,
  type ImageSize,
  type PixelRect
} from './image-region.ts'
import { measureSlideImage } from './image-screen-rect.ts'
import { selectedSlideImage } from './panel-selection.ts'

const GREEN = '#16a34a'
const PRIMARY = '#14A6EA'
const INK = '#00141A'
const LINE = 'rgba(0, 20, 26, 0.08)'
const BUBBLE_SHADOW =
  '0px 9px 28px 8px rgba(0, 20, 26, 0.05), 0px 6px 16px 0px rgba(0, 0, 0, 0.08), 0px 3px 6px -4px rgba(0, 0, 0, 0.12)'
const INPUT_WIDTH = 280
const FOLLOW_SELECTION_MS = 200

interface Copy {
  readonly edit: string
  readonly back: string
  readonly hint: string
  readonly apply: (count: number) => string
  readonly placeholder: string
  readonly generating: string
  readonly stop: string
  readonly background: string
  readonly max: string
  readonly exitTitle: string
  readonly exitBody: string
  readonly keepEditing: string
  readonly confirmExit: string
  readonly failed: string
  readonly loading: string
  readonly delete: string
}

interface Annotation {
  readonly pageId: string
  readonly elementId: string
  readonly regions: readonly ImageRegion[]
  readonly baseline: ImageRegion
  readonly selectedId: string
}

interface SlideImageSelection {
  readonly pageId: string
  readonly elementId: string
}

type Phase = 'selected' | 'editing' | 'generating'

type Gesture =
  | {
      readonly type: 'create'
      startX: number
      startY: number
      x: number
      y: number
    }
  | {
      readonly type: 'move'
      readonly id: string
      readonly origin: ImageRegion
      readonly clientX: number
      readonly clientY: number
    }
  | {
      readonly type: 'resize'
      readonly id: string
      readonly handle: Handle
      readonly origin: ImageRegion
      readonly clientX: number
      readonly clientY: number
    }

/**
 * Mount the region editor for one editable embedded Slide.
 *
 * The overlay lives on the Viewer document, which is the iframe that owns the canvas. A standalone
 * or read-only address has nothing to apply an edit to, so it mounts nothing.
 */
export function mountImageRegionEditor(univer: Univer, unitId: string): () => void {
  const location = readLocation()
  if (location.mode !== 'embedded' || location.comparison !== null) return () => undefined
  const sessionId = new URLSearchParams(window.location.search).get('sessionId')
  if (sessionId === null || sessionId.length === 0) return () => undefined
  const host = document.createElement('div')
  host.dataset.imageRegionRoot = 'true'
  host.style.position = 'fixed'
  host.style.inset = '0'
  host.style.pointerEvents = 'none'
  host.style.zIndex = '60'
  document.body.append(host)
  const stopWheel = forwardSlideZoom(host, univer, unitId)
  const root = createRoot(host)
  flushSync(() => {
    root.render(<ImageRegionEditor univer={univer} unitId={unitId} sessionId={sessionId} />)
  })
  return () => {
    stopWheel()
    root.unmount()
    host.remove()
  }
}

/**
 * Ctrl/trackpad zoom must change the slide, not the browser page.
 *
 * The overlay sits above the canvas, so a wheel there never reaches the slide and the browser
 * scales the input text instead. Apply the same zoom step the canvas uses, and keep this
 * document's own zoom at 1 so the input stays a fixed size.
 */
function forwardSlideZoom(host: HTMLElement, univer: Univer, unitId: string): () => void {
  const onWheel = (event: WheelEvent): void => {
    if (!event.ctrlKey && !event.metaKey) return
    event.preventDefault()
    try {
      const injector = univer.__getInjector()
      const unit = injector
        .get(IUniverInstanceService)
        .getUnit<SlideModel>(unitId, UniverInstanceType.UNIVER_SLIDE)
      if (unit === null || unit === undefined) return
      const current = unit.getZoomRatio()
      const next = getNextWheelZoomRatio(current, event)
      if (next === current) return
      void injector
        .get(ICommandService)
        .executeCommand(
          SetSlideZoomRatioOperation.id,
          { unitId, zoomRatio: next },
          { onlyLocal: true }
        )
    } catch {
      // The slide unit can already be disposed while this overlay is still mounted.
    }
  }
  host.addEventListener('wheel', onWheel, { capture: true, passive: false })
  return () => host.removeEventListener('wheel', onWheel, { capture: true })
}

function ImageRegionEditor({
  univer,
  unitId,
  sessionId
}: {
  univer: Univer
  unitId: string
  sessionId: string
}): ReactElement {
  const copy = text()
  const [selection, setSelection] = useState<SlideImageSelection | null>(null)
  const [session, setSession] = useState<Annotation | null>(null)
  const [phase, setPhase] = useState<Phase>('selected')
  const [rect, setRect] = useState<PixelRect | null>(null)
  const [bitmap, setBitmap] = useState<ImageSize | null>(null)
  const [preview, setPreview] = useState<ImageRegion | null>(null)
  const [toast, setToast] = useState<string | null>(null)
  const [confirm, setConfirm] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [slow, setSlow] = useState(false)
  const draftsRef = useRef(new Map<string, Annotation>())
  const sessionRef = useRef<Annotation | null>(null)
  const rectRef = useRef<PixelRect | null>(null)
  const bitmapRef = useRef<ImageSize | null>(null)
  const phaseRef = useRef<Phase>('selected')
  const selectionRef = useRef<SlideImageSelection | null>(null)
  const gestureRef = useRef<Gesture | null>(null)
  const abortRef = useRef<AbortController | null>(null)
  const slowTimerRef = useRef<number | undefined>(undefined)
  const toastTimerRef = useRef<number | undefined>(undefined)
  const mountedRef = useRef(true)

  useLayoutEffect(() => {
    sessionRef.current = session
    rectRef.current = rect
    bitmapRef.current = bitmap
    phaseRef.current = phase
    selectionRef.current = selection
  }, [session, rect, bitmap, phase, selection])

  useEffect(() => {
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      abortRef.current?.abort()
      if (slowTimerRef.current !== undefined) window.clearTimeout(slowTimerRef.current)
      if (toastTimerRef.current !== undefined) window.clearTimeout(toastTimerRef.current)
    }
  }, [])

  useEffect(() => {
    let subscription: { unsubscribe(): void } | undefined
    let lifecycle: { unsubscribe(): void } | undefined
    let cancelled = false
    try {
      const injector = univer.__getInjector()
      const start = (): void => {
        if (cancelled) return
        try {
          const drawings = injector.get(ISlideDrawingService)
          const state = injector.get(ISlideDrawingStateService)
          const applySelection = (snapshot: ISlideDrawingStateSnapshot): void => {
            setSelection(selectedSlideImage(drawings, unitId, snapshot))
          }
          // state$ does not replay, and the image may already be selected when this overlay mounts.
          applySelection(state.getSnapshot())
          subscription = state.state$.subscribe(applySelection)
        } catch {
          // A build without the Slide UI plugins has no drawing selection to annotate.
        }
      }
      lifecycle = injector
        .get(LifecycleService)
        .lifecycle$.pipe(
          filter((stage) => stage >= LifecycleStages.Rendered),
          take(1)
        )
        .subscribe(start)
    } catch {
      // The viewer can dispose the injector before the first subscription attaches.
      return () => undefined
    }
    return () => {
      cancelled = true
      lifecycle?.unsubscribe()
      subscription?.unsubscribe()
    }
  }, [univer, unitId])

  const pageId = phase === 'selected' ? selection?.pageId : session?.pageId
  const elementId = phase === 'selected' ? selection?.elementId : session?.elementId
  useEffect(() => {
    let frame = 0
    const update = (): void => {
      if (pageId === undefined || elementId === undefined) {
        setRect((current) => (current === null ? current : null))
        setBitmap((current) => (current === null ? current : null))
        return
      }
      const measured = measureSlideImage(univer, unitId, { pageId, elementId })
      // A zoom step can rebuild the drawing for one frame. Keep the last rectangle so the boxes
      // do not vanish and pop back somewhere else.
      if (measured === null) return
      setRect((current) => (sameRect(current, measured.rect) ? current : measured.rect))
      setBitmap((current) => (sameSize(current, measured.bitmap) ? current : measured.bitmap))
    }
    const follow = (): void => {
      update()
      frame = window.requestAnimationFrame(follow)
    }
    // Pan and zoom change the canvas without a React event, so follow them on every frame.
    follow()
    return () => window.cancelAnimationFrame(frame)
  }, [pageId, elementId, univer, unitId])

  const updateSession = useCallback((next: Annotation) => {
    draftsRef.current.set(annotationKey(next), next)
    sessionRef.current = next
    setSession(next)
  }, [])

  const leave = useCallback((drop: boolean) => {
    const current = sessionRef.current
    if (current !== null && drop) draftsRef.current.delete(annotationKey(current))
    sessionRef.current = null
    setSession(null)
    setConfirm(false)
    setError(null)
    setPreview(null)
    gestureRef.current = null
    setPhase('selected')
  }, [])

  useEffect(() => {
    if (phase !== 'editing') return
    const current = sessionRef.current
    if (current === null || isDirty(current.regions, current.baseline)) return
    if (
      selection !== null &&
      selection.pageId === current.pageId &&
      selection.elementId === current.elementId
    ) {
      return
    }
    const timer = window.setTimeout(() => {
      const latest = sessionRef.current
      const selected = selectionRef.current
      if (
        phaseRef.current !== 'editing' ||
        latest === null ||
        isDirty(latest.regions, latest.baseline)
      ) {
        return
      }
      if (
        selected !== null &&
        selected.pageId === latest.pageId &&
        selected.elementId === latest.elementId
      ) {
        return
      }
      leave(true)
    }, FOLLOW_SELECTION_MS)
    return () => window.clearTimeout(timer)
  }, [phase, selection, leave])

  const selectedRegionId = session?.selectedId
  useLayoutEffect(() => {
    if (phase !== 'editing' || selectedRegionId === undefined) return
    const node = document.querySelector<HTMLTextAreaElement>(
      `textarea[data-region-id="${CSS.escape(selectedRegionId)}"]`
    )
    node?.focus({ preventScroll: true })
  }, [phase, selectedRegionId])

  const showToast = useCallback((message: string) => {
    setToast(message)
    if (toastTimerRef.current !== undefined) window.clearTimeout(toastTimerRef.current)
    toastTimerRef.current = window.setTimeout(() => setToast(null), 2600)
  }, [])

  const begin = (): void => {
    if (selection === null || rect === null) return
    const key = `${selection.pageId}:${selection.elementId}`
    const saved = draftsRef.current.get(key)
    const region = createDefaultRegion({ width: rect.width, height: rect.height })
    const next: Annotation = saved ?? {
      pageId: selection.pageId,
      elementId: selection.elementId,
      regions: [region],
      baseline: region,
      selectedId: region.id
    }
    updateSession(next)
    setError(null)
    setPhase('editing')
  }

  const onBack = (): void => {
    const current = sessionRef.current
    if (current !== null && isDirty(current.regions, current.baseline)) {
      setConfirm(true)
      return
    }
    leave(true)
  }

  const stop = (): void => {
    abortRef.current?.abort()
    setSlow(false)
    setPhase('selected')
    sessionRef.current = null
    setSession(null)
    setPreview(null)
  }

  const apply = (): void => {
    const current = sessionRef.current
    const pixels = bitmapRef.current
    if (current === null || pixels === null) return
    const prompt = promptOf(current.regions, pixels)
    if (prompt.length === 0) return
    const controller = new AbortController()
    abortRef.current = controller
    setPhase('generating')
    setError(null)
    setSlow(false)
    setPreview(null)
    if (slowTimerRef.current !== undefined) window.clearTimeout(slowTimerRef.current)
    slowTimerRef.current = window.setTimeout(() => {
      if (mountedRef.current) setSlow(true)
    }, 3000)
    const key = annotationKey(current)
    void submitEdit(sessionId, current, prompt, controller).then(
      (result) => {
        if (!mountedRef.current || controller.signal.aborted) return undefined
        if (slowTimerRef.current !== undefined) window.clearTimeout(slowTimerRef.current)
        if (abortRef.current === controller) abortRef.current = null
        if (!result.ok) {
          setError(result.message ?? copy.failed)
          setPhase('editing')
          return undefined
        }
        draftsRef.current.delete(key)
        sessionRef.current = null
        setSession(null)
        setPhase('selected')
        return undefined
      },
      (caught: unknown) => {
        if (!mountedRef.current || controller.signal.aborted || isAbortError(caught))
          return undefined
        if (slowTimerRef.current !== undefined) window.clearTimeout(slowTimerRef.current)
        if (abortRef.current === controller) abortRef.current = null
        setError(
          caught instanceof Error && caught.message.length > 0 ? caught.message : copy.failed
        )
        setPhase('editing')
        return undefined
      }
    )
  }

  const display = (): ImageSize => {
    const current = rectRef.current
    return { width: current?.width ?? 1, height: current?.height ?? 1 }
  }

  const onCreateDown = (event: ReactPointerEvent<HTMLDivElement>): void => {
    const image = rectRef.current
    if (image === null || phaseRef.current !== 'editing') return
    event.preventDefault()
    capture(event.currentTarget, event.pointerId)
    const local = imageLocal(event, image)
    const x = local.x / image.width
    const y = local.y / image.height
    gestureRef.current = { type: 'create', startX: x, startY: y, x, y }
  }

  const onRegionDown = (event: ReactPointerEvent<HTMLDivElement>, region: ImageRegion): void => {
    const current = sessionRef.current
    if (current === null) return
    event.preventDefault()
    event.stopPropagation()
    capture(event.currentTarget, event.pointerId)
    if (current.selectedId !== region.id) updateSession({ ...current, selectedId: region.id })
    gestureRef.current = {
      type: 'move',
      id: region.id,
      origin: region,
      clientX: event.clientX,
      clientY: event.clientY
    }
  }

  const onHandleDown = (
    event: ReactPointerEvent<HTMLSpanElement>,
    region: ImageRegion,
    handle: Handle
  ): void => {
    event.preventDefault()
    event.stopPropagation()
    capture(event.currentTarget, event.pointerId)
    gestureRef.current = {
      type: 'resize',
      id: region.id,
      handle,
      origin: region,
      clientX: event.clientX,
      clientY: event.clientY
    }
  }

  const onGestureMove = (event: ReactPointerEvent<HTMLElement>): void => {
    const gesture = gestureRef.current
    const image = rectRef.current
    const current = sessionRef.current
    if (gesture === null || image === null || current === null) return
    if (gesture.type === 'create') {
      const local = imageLocal(event, image)
      gesture.x = local.x / image.width
      gesture.y = local.y / image.height
      const dx = (gesture.x - gesture.startX) * image.width
      const dy = (gesture.y - gesture.startY) * image.height
      setPreview(
        isDrag(dx, dy)
          ? createDragRegion(
              display(),
              gesture.startX * image.width,
              gesture.startY * image.height,
              local.x,
              local.y,
              'preview'
            )
          : null
      )
      return
    }
    const dx = event.clientX - gesture.clientX
    const dy = event.clientY - gesture.clientY
    const next =
      gesture.type === 'move'
        ? moveRegion(gesture.origin, dx, dy, display())
        : resizeRegion(gesture.origin, gesture.handle, dx, dy, display())
    updateSession({
      ...current,
      regions: current.regions.map((region) => (region.id === gesture.id ? next : region))
    })
  }

  const onGestureEnd = (): void => {
    const gesture = gestureRef.current
    gestureRef.current = null
    setPreview(null)
    if (gesture === null || gesture.type !== 'create') return
    const current = sessionRef.current
    const image = rectRef.current
    if (current === null || image === null) return
    if (!canAddRegion(current.regions)) {
      showToast(copy.max)
      return
    }
    const dx = (gesture.x - gesture.startX) * image.width
    const dy = (gesture.y - gesture.startY) * image.height
    const id = nextRegionId(current.regions)
    const created = isDrag(dx, dy)
      ? createDragRegion(
          display(),
          gesture.startX * image.width,
          gesture.startY * image.height,
          gesture.x * image.width,
          gesture.y * image.height,
          id
        )
      : createPointRegion(
          display(),
          gesture.startX * image.width,
          gesture.startY * image.height,
          id
        )
    updateSession({
      ...current,
      regions: [...current.regions, created],
      selectedId: created.id
    })
  }

  const setText = (id: string, value: string): void => {
    const current = sessionRef.current
    if (current === null) return
    updateSession({
      ...current,
      regions: current.regions.map((region) =>
        region.id === id ? withText(region, value) : region
      )
    })
  }

  const removeRegion = (id: string): void => {
    const current = sessionRef.current
    if (current === null || current.regions.length <= 1) return
    const regions = current.regions.filter((region) => region.id !== id)
    const selected = regions.some((region) => region.id === current.selectedId)
      ? current.selectedId
      : (regions[regions.length - 1]?.id ?? current.selectedId)
    updateSession({ ...current, regions, selectedId: selected })
  }

  const buttonAt = rect === null ? null : placeButton(rect)
  const editing = phase === 'editing' && session !== null
  const generating = phase === 'generating'

  return (
    <>
      <style>{REGION_STYLE}</style>
      {generating && <div style={blockerStyle} />}
      {rect !== null && (phase === 'selected' || editing || generating) && (
        <div data-image-region-frame="true" style={frameStyle(rect)} />
      )}
      {generating && rect !== null && (
        <div style={blurFrameStyle(rect)}>
          <div style={blurStyle} />
          <div style={sweepStyle} />
        </div>
      )}
      {editing && rect !== null && session !== null && (
        <div
          data-image-region-hit="true"
          style={hitStyle(rect)}
          onPointerDown={onCreateDown}
          onPointerMove={onGestureMove}
          onPointerUp={onGestureEnd}
          onPointerCancel={onGestureEnd}
        />
      )}
      {editing && rect !== null && session !== null && preview !== null && (
        <div style={regionStyle(regionFrame(preview, rect), false, 8)} />
      )}
      {editing &&
        rect !== null &&
        session !== null &&
        session.regions.map((region, index) => {
          const box = regionFrame(region, rect)
          const selected = region.id === session.selectedId
          return (
            <div
              key={region.id}
              data-image-region="true"
              data-image-region-kind={region.kind}
              style={regionStyle(box, selected, 12 + index)}
              onPointerDown={(event) => onRegionDown(event, region)}
              onPointerMove={onGestureMove}
              onPointerUp={onGestureEnd}
              onPointerCancel={onGestureEnd}
            >
              {selected &&
                (['nw', 'ne', 'sw', 'se'] as const).map((handle) => (
                  <span
                    key={handle}
                    data-image-region-handle={handle}
                    style={handleStyle(handle)}
                    onPointerDown={(event) => onHandleDown(event, region, handle)}
                    onPointerMove={onGestureMove}
                    onPointerUp={onGestureEnd}
                    onPointerCancel={onGestureEnd}
                  />
                ))}
            </div>
          )
        })}
      {editing &&
        rect !== null &&
        session !== null &&
        session.regions.map((region, index) => {
          const box = regionFrame(region, rect)
          const input = placeInput(box, rect)
          const selected = region.id === session.selectedId
          return (
            <div
              key={`${region.id}-input`}
              style={inputStyle(input, selected, 40 + index)}
              onPointerDown={(event) => event.stopPropagation()}
            >
              {selected && <Caret above={input.above} />}
              <span style={badgeStyle(selected)}>{index + 1}</span>
              <textarea
                data-region-id={region.id}
                value={region.text}
                placeholder={copy.placeholder}
                rows={1}
                style={textareaStyle}
                ref={(node) => {
                  if (node !== null) fitTextarea(node)
                }}
                onChange={(event) => {
                  setText(region.id, event.target.value)
                  fitTextarea(event.target)
                }}
                onFocus={() => {
                  const current = sessionRef.current
                  if (current === null || current.selectedId === region.id) return
                  updateSession({ ...current, selectedId: region.id })
                }}
              />
              {session.regions.length > 1 && (
                <button
                  type="button"
                  data-image-region-action="delete"
                  aria-label={copy.delete}
                  style={deleteStyle}
                  onClick={() => removeRegion(region.id)}
                >
                  <TrashIcon />
                </button>
              )}
            </div>
          )
        })}
      {phase === 'selected' && selection !== null && buttonAt !== null && (
        <button
          type="button"
          data-image-region-action="edit"
          style={editButtonStyle(buttonAt)}
          onClick={begin}
        >
          {copy.edit}
        </button>
      )}
      {editing && session !== null && (
        <div
          data-image-region-bar="true"
          style={rect === null ? viewportBarStyle : imageDockStyle(placeButton(rect))}
        >
          <button
            type="button"
            data-image-region-action="back"
            aria-label={copy.back}
            style={backButtonStyle}
            onClick={onBack}
          >
            <BackIcon />
          </button>
          <span style={hintStyle}>{copy.hint}</span>
          <button
            type="button"
            data-image-region-action="apply"
            style={applyButtonStyle(canApply(session.regions) && bitmap !== null)}
            disabled={!canApply(session.regions) || bitmap === null}
            {...(bitmap === null ? { title: copy.loading } : {})}
            onClick={apply}
          >
            {copy.apply(session.regions.length)}
          </button>
        </div>
      )}
      {generating && (
        <div data-image-region-bar="true" style={viewportBarStyle}>
          <span style={spinnerStyle} />
          <span style={generatingTextStyle}>
            {copy.generating}
            {slow && <span style={slowStyle}>{copy.background}</span>}
          </span>
          <button
            type="button"
            data-image-region-action="stop"
            style={quietButtonStyle}
            onClick={stop}
          >
            {copy.stop}
          </button>
        </div>
      )}
      {error !== null && phase === 'editing' && (
        <div role="alert" style={errorStyle}>
          {error}
        </div>
      )}
      {toast !== null && <output style={toastStyle}>{toast}</output>}
      {confirm && (
        <div style={confirmBackdropStyle}>
          <div role="alertdialog" aria-modal="true" style={confirmStyle}>
            <p style={confirmTitleStyle}>{copy.exitTitle}</p>
            <p style={confirmBodyStyle}>{copy.exitBody}</p>
            <div style={confirmActionsStyle}>
              <button type="button" style={quietButtonStyle} onClick={() => setConfirm(false)}>
                {copy.keepEditing}
              </button>
              <button
                type="button"
                data-image-region-action="confirm-exit"
                style={dangerButtonStyle}
                onClick={() => leave(true)}
              >
                {copy.confirmExit}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  )
}

const COPY: Record<'zh' | 'en', Copy> = {
  zh: {
    edit: '编辑区域',
    back: '返回',
    hint: '拖拽或点击添加区域',
    apply: (count) => `全部应用 (${count})`,
    placeholder: '描述要进行的修改…',
    generating: '正在编辑图片…',
    stop: '停止',
    background: '任务正在后台处理，你可以先去忙别的。',
    max: '每次最多编辑8个区域',
    exitTitle: '确定退出编辑吗？',
    exitBody: '尚未应用的区域和修改内容将丢失。',
    keepEditing: '继续编辑',
    confirmExit: '确认退出',
    failed: '生成失败',
    loading: '图片仍在加载',
    delete: '删除'
  },
  en: {
    edit: 'Edit region',
    back: 'Back',
    hint: 'Drag or click to add a region',
    apply: (count) => `Apply all (${count})`,
    placeholder: 'Describe the change…',
    generating: 'Editing image…',
    stop: 'Stop',
    background: 'This is running in the background. You can do something else.',
    max: 'You can edit at most 8 regions at a time',
    exitTitle: 'Leave without applying?',
    exitBody: 'Regions and text that have not been applied will be lost.',
    keepEditing: 'Keep editing',
    confirmExit: 'Leave',
    failed: 'Could not edit the image',
    loading: 'The image is still loading',
    delete: 'Delete'
  }
}

function text(): Copy {
  return currentLang().startsWith('zh') ? COPY.zh : COPY.en
}

async function submitEdit(
  sessionId: string,
  session: Annotation,
  prompt: string,
  controller: AbortController
): Promise<{ ok: true } | { ok: false; message?: string }> {
  const response = await fetch('/univer-api/panel/image-edit', {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'content-type': 'application/json' },
    signal: controller.signal,
    body: JSON.stringify({
      sessionId,
      pageId: session.pageId,
      elementId: session.elementId,
      prompt
    })
  })
  const body = await response.text()
  if (controller.signal.aborted) return { ok: false }
  let payload: { ok?: unknown; message?: unknown } = {}
  if (body.length > 0) {
    try {
      payload = JSON.parse(body) as { ok?: unknown; message?: unknown }
    } catch {
      // A non-JSON failure still has to stay on the editor instead of replacing the image.
      return { ok: false, message: text().failed }
    }
  }
  if (response.ok && payload.ok === true) return { ok: true }
  return {
    ok: false,
    ...(typeof payload.message === 'string' && payload.message.length > 0
      ? { message: payload.message }
      : { message: text().failed })
  }
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError'
}

function withText(region: ImageRegion, value: string): ImageRegion {
  return {
    id: region.id,
    kind: region.kind,
    x: region.x,
    y: region.y,
    w: region.w,
    h: region.h,
    text: value
  }
}

function annotationKey(session: Annotation): string {
  return `${session.pageId}:${session.elementId}`
}

function imageLocal(
  event: { clientX: number; clientY: number },
  image: PixelRect
): { x: number; y: number } {
  return { x: event.clientX - image.left, y: event.clientY - image.top }
}

function capture(element: HTMLElement, pointerId: number): void {
  try {
    element.setPointerCapture(pointerId)
  } catch {
    // The pointer can already be released before capture is requested.
  }
}

function fitTextarea(node: HTMLTextAreaElement): void {
  node.style.height = 'auto'
  node.style.height = `${Math.min(node.scrollHeight, 72)}px`
}

function sameRect(current: PixelRect | null, next: PixelRect | null): boolean {
  if (current === null || next === null) return current === next
  return (
    current.left === next.left &&
    current.top === next.top &&
    current.width === next.width &&
    current.height === next.height
  )
}

function sameSize(current: ImageSize | null, next: ImageSize | null): boolean {
  if (current === null || next === null) return current === next
  return current.width === next.width && current.height === next.height
}

function placeButton(rect: PixelRect): { left: number; top: number } {
  // Stay on the image's bottom edge while it zooms. Moving the button above the image makes it
  // look detached from the selection.
  return { left: rect.left + rect.width / 2, top: rect.top + rect.height + 4 }
}

function placeInput(
  box: PixelRect,
  image: PixelRect
): { left: number; top: number; width: number; above: boolean } {
  const left = Math.min(Math.max(8, box.left), Math.max(8, window.innerWidth - INPUT_WIDTH - 8))
  const below = box.top + box.height + 8
  // The edit dock occupies the strip under the image, so a bubble that would land there sits
  // above its region instead.
  const dockTop = image.top + image.height + 4
  const above = below + 44 > dockTop && box.top > 56
  const top = above ? Math.max(8, box.top - 52) : below
  return { left, top, width: INPUT_WIDTH, above }
}

function frameStyle(rect: PixelRect): CSSProperties {
  return {
    position: 'fixed',
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
    outline: `2px solid ${GREEN}`,
    pointerEvents: 'none',
    boxSizing: 'border-box',
    zIndex: 1
  }
}

function hitStyle(rect: PixelRect): CSSProperties {
  return {
    ...frameStyle(rect),
    outline: 'none',
    pointerEvents: 'auto',
    cursor: 'crosshair',
    touchAction: 'none',
    zIndex: 2
  }
}

function regionStyle(box: PixelRect, selected: boolean, zIndex: number): CSSProperties {
  return {
    position: 'fixed',
    left: box.left,
    top: box.top,
    width: box.width,
    height: box.height,
    border: `2px solid ${GREEN}`,
    background: selected ? 'rgba(22, 163, 74, 0.16)' : 'rgba(22, 163, 74, 0.08)',
    boxSizing: 'border-box',
    pointerEvents: 'auto',
    cursor: 'move',
    touchAction: 'none',
    zIndex
  }
}

function handleStyle(handle: Handle): CSSProperties {
  const style: CSSProperties = {
    position: 'absolute',
    width: 10,
    height: 10,
    borderRadius: 999,
    background: '#fff',
    border: `2px solid ${GREEN}`,
    boxSizing: 'border-box',
    pointerEvents: 'auto',
    cursor: handle === 'nw' || handle === 'se' ? 'nwse-resize' : 'nesw-resize',
    touchAction: 'none'
  }
  if (handle === 'nw' || handle === 'sw') style.left = -5
  else style.right = -5
  if (handle === 'nw' || handle === 'ne') style.top = -5
  else style.bottom = -5
  return style
}

function inputStyle(
  input: { left: number; top: number; width: number },
  selected: boolean,
  zIndex: number
): CSSProperties {
  return {
    position: 'fixed',
    left: input.left,
    top: input.top,
    width: input.width,
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    padding: 8,
    background: '#fff',
    border: `1px solid ${selected ? PRIMARY : LINE}`,
    borderRadius: 12,
    boxShadow: selected ? BUBBLE_SHADOW : 'none',
    pointerEvents: 'auto',
    zIndex,
    boxSizing: 'border-box'
  }
}

function Caret(props: { readonly above: boolean }): ReactElement {
  return (
    <>
      <span style={caretStyle(props.above, false)} />
      <span style={caretStyle(props.above, true)} />
    </>
  )
}

function caretStyle(above: boolean, fill: boolean): CSSProperties {
  const shared: CSSProperties = {
    position: 'absolute',
    left: '50%',
    width: 0,
    height: 0,
    transform: 'translateX(-50%)',
    pointerEvents: 'none'
  }
  const size = fill ? 7 : 8
  const edge = `${size}px solid transparent`
  const color = fill ? '#fff' : PRIMARY
  if (above) {
    return {
      ...shared,
      bottom: fill ? -6 : -8,
      borderLeft: edge,
      borderRight: edge,
      borderTop: `${size}px solid ${color}`
    }
  }
  return {
    ...shared,
    top: fill ? -6 : -8,
    borderLeft: edge,
    borderRight: edge,
    borderBottom: `${size}px solid ${color}`
  }
}

function editButtonStyle(at: { left: number; top: number }): CSSProperties {
  return {
    ...imageDockStyle(at),
    height: 44,
    padding: '8px 16px',
    color: 'rgba(0, 20, 26, 0.85)',
    fontSize: 14,
    lineHeight: '22px',
    fontWeight: 400,
    cursor: 'pointer',
    zIndex: 5
  }
}

function imageDockStyle(at: { left: number; top: number }): CSSProperties {
  return {
    position: 'fixed',
    left: at.left,
    top: at.top,
    transform: 'translateX(-50%)',
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    maxWidth: 'calc(100% - 16px)',
    height: 44,
    padding: '0 8px',
    boxSizing: 'border-box',
    background: '#fff',
    border: `1px solid ${LINE}`,
    borderRadius: 8,
    boxShadow: BUBBLE_SHADOW,
    pointerEvents: 'auto',
    zIndex: 80,
    whiteSpace: 'nowrap'
  }
}

const blockerStyle: CSSProperties = {
  position: 'fixed',
  inset: 0,
  pointerEvents: 'auto',
  zIndex: 3
}

const blurStyle: CSSProperties = {
  position: 'absolute',
  inset: 0,
  backdropFilter: 'blur(8px)',
  background: 'rgba(255, 255, 255, 0.28)'
}

const sweepStyle: CSSProperties = {
  position: 'absolute',
  top: 0,
  bottom: 0,
  width: '45%',
  background: 'linear-gradient(90deg, transparent, rgba(22, 163, 74, 0.45), transparent)',
  animation: 'uv-region-sweep 1.4s linear infinite'
}

function blurFrameStyle(rect: PixelRect): CSSProperties {
  return {
    position: 'fixed',
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
    overflow: 'hidden',
    pointerEvents: 'none',
    zIndex: 4
  }
}

const viewportBarStyle: CSSProperties = {
  position: 'fixed',
  left: '50%',
  bottom: 16,
  transform: 'translateX(-50%)',
  display: 'flex',
  alignItems: 'center',
  gap: 12,
  maxWidth: 'calc(100% - 24px)',
  padding: '8px 12px',
  background: '#fff',
  border: `1px solid ${LINE}`,
  borderRadius: 8,
  boxShadow: BUBBLE_SHADOW,
  pointerEvents: 'auto',
  zIndex: 80
}

const backButtonStyle: CSSProperties = {
  width: 24,
  height: 24,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  padding: 0,
  border: 'none',
  borderRadius: 4,
  background: 'rgba(0, 20, 26, 0.02)',
  color: 'rgba(0, 20, 26, 0.6)',
  cursor: 'pointer',
  flexShrink: 0
}

const quietButtonStyle: CSSProperties = {
  height: 32,
  padding: '0 12px',
  borderRadius: 8,
  border: '1px solid #e5e5e5',
  background: '#fff',
  color: '#171717',
  font: '13px/30px inherit',
  cursor: 'pointer'
}

function applyButtonStyle(enabled: boolean): CSSProperties {
  return {
    height: 28,
    padding: '4px 12px',
    borderRadius: 16,
    border: 'none',
    background: INK,
    color: '#fff',
    fontSize: 12,
    lineHeight: '20px',
    fontWeight: 500,
    cursor: enabled ? 'pointer' : 'default',
    opacity: enabled ? 1 : 0.45,
    flexShrink: 0
  }
}

const dangerButtonStyle: CSSProperties = {
  ...quietButtonStyle,
  border: 'none',
  background: '#dc2626',
  color: '#fff'
}

const hintStyle: CSSProperties = {
  color: 'rgba(0, 20, 26, 0.45)',
  fontSize: 12,
  lineHeight: '20px',
  fontWeight: 400,
  whiteSpace: 'nowrap'
}

function badgeStyle(selected: boolean): CSSProperties {
  return {
    width: 20,
    height: 20,
    borderRadius: 8,
    background: selected ? INK : 'rgba(0, 20, 26, 0.45)',
    color: '#fff',
    fontSize: 12,
    lineHeight: '20px',
    fontWeight: 500,
    textAlign: 'center',
    flexShrink: 0
  }
}

const textareaStyle: CSSProperties = {
  flex: 1,
  minWidth: 0,
  border: 'none',
  outline: 'none',
  resize: 'none',
  padding: '0 8px 0 2px',
  margin: 0,
  fontSize: 14,
  lineHeight: '24px',
  fontWeight: 400,
  fontFamily: 'inherit',
  color: INK,
  background: 'transparent',
  maxHeight: 72,
  overflow: 'auto'
}

const deleteStyle: CSSProperties = {
  width: 24,
  height: 24,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  color: INK,
  opacity: 0.5,
  cursor: 'pointer',
  padding: 0,
  flexShrink: 0
}

function BackIcon(): ReactElement {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path
        d="M10.2 3.2 5.4 8l4.8 4.8"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function TrashIcon(): ReactElement {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path
        d="M3.2 4.4h9.6M6.2 4.4V3.2h3.6v1.2M4.4 4.4l.6 8.2h5.9l.6-8.2"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

const generatingTextStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 2,
  color: '#171717',
  fontSize: 13
}

const slowStyle: CSSProperties = {
  color: '#737373',
  fontSize: 12
}

const spinnerStyle: CSSProperties = {
  width: 14,
  height: 14,
  borderRadius: 999,
  border: '2px solid #e5e5e5',
  borderTopColor: GREEN,
  animation: 'uv-region-spin 0.8s linear infinite',
  flexShrink: 0
}

const errorStyle: CSSProperties = {
  position: 'fixed',
  left: '50%',
  bottom: 68,
  transform: 'translateX(-50%)',
  maxWidth: 'calc(100% - 24px)',
  padding: '6px 10px',
  borderRadius: 8,
  background: '#fef2f2',
  color: '#b91c1c',
  fontSize: 13,
  pointerEvents: 'none',
  zIndex: 81
}

const toastStyle: CSSProperties = {
  display: 'block',
  position: 'fixed',
  left: '50%',
  top: 16,
  transform: 'translateX(-50%)',
  padding: '8px 12px',
  borderRadius: 8,
  background: 'rgba(23, 23, 23, 0.92)',
  color: '#fff',
  fontSize: 13,
  pointerEvents: 'none',
  zIndex: 90
}

const confirmBackdropStyle: CSSProperties = {
  position: 'fixed',
  inset: 0,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  background: 'rgba(0, 0, 0, 0.28)',
  pointerEvents: 'auto',
  zIndex: 100
}

const confirmStyle: CSSProperties = {
  width: 320,
  maxWidth: 'calc(100% - 32px)',
  padding: 16,
  borderRadius: 12,
  background: '#fff',
  boxShadow: '0 16px 40px rgba(0, 0, 0, 0.18)'
}

const confirmTitleStyle: CSSProperties = {
  margin: 0,
  color: '#171717',
  fontSize: 15,
  fontWeight: 600
}

const confirmBodyStyle: CSSProperties = {
  margin: '8px 0 0',
  color: '#525252',
  fontSize: 13,
  lineHeight: '20px'
}

const confirmActionsStyle: CSSProperties = {
  display: 'flex',
  justifyContent: 'flex-end',
  gap: 8,
  marginTop: 16
}

const REGION_STYLE = `
@keyframes uv-region-spin { to { transform: rotate(360deg); } }
@keyframes uv-region-sweep { from { transform: translateX(-120%); } to { transform: translateX(280%); } }
`
