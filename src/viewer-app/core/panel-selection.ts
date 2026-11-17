import type { Univer } from '@univerjs/core'
import { ISlideDrawingService, PageElementTypeEnum } from '@univerjs-pro/slides'
import { ISlideDrawingStateService, type ISlideDrawingStateSnapshot } from '@univerjs-pro/slides-ui'
import { readLocation } from './config'

/** The image one Viewer reports as selected, in the shape the Host route accepts. */
interface PanelSelection {
  readonly pageId: string
  readonly elementId: string
}

/** Delay applied to selection reports so hover and drag churn never reaches the Host. */
const REPORT_DEBOUNCE_MS = 120

/**
 * Report the image the user has selected in the right panel to the Host.
 *
 * The Host cannot see browser state, yet "replace this image" is meaningless without it, so the
 * selection travels as a small structured report. Only an embedded Viewer bound to a DSH session
 * reports anything: a standalone address has no conversation to attach the selection to.
 */
export function watchPanelSelection(univer: Univer, unitId: string): () => void {
  const loc = readLocation()
  if (loc.mode !== 'embedded') return () => undefined
  const sessionId = new URLSearchParams(location.search).get('sessionId')
  if (sessionId === null || sessionId.length === 0) return () => undefined

  let lastKey = ''
  let timer: ReturnType<typeof setTimeout> | undefined
  const publish = (selection: PanelSelection | null): void => {
    const key = selection === null ? '' : `${selection.pageId}:${selection.elementId}`
    if (key === lastKey) return
    lastKey = key
    const body =
      selection === null
        ? { sessionId }
        : { sessionId, pageId: selection.pageId, elementId: selection.elementId }
    void postSelection(body)
  }

  let drawings: ISlideDrawingService
  let subscription: { unsubscribe(): void }
  try {
    const injector = univer.__getInjector()
    const state = injector.get(ISlideDrawingStateService)
    drawings = injector.get(ISlideDrawingService)
    subscription = state.state$.subscribe((snapshot) => {
      if (timer !== undefined) clearTimeout(timer)
      timer = setTimeout(() => {
        timer = undefined
        publish(selectedImage(drawings, unitId, snapshot))
      }, REPORT_DEBOUNCE_MS)
    })
  } catch {
    // A build without the Slide UI plugins registers no drawing selection to report.
    return () => undefined
  }

  return () => {
    if (timer !== undefined) clearTimeout(timer)
    subscription.unsubscribe()
  }
}

/**
 * The selected image element, or `null` when the selection is anything else.
 *
 * Text, shapes, and multi-selections are not image targets, so they clear the Host's selection
 * rather than leaving a stale image behind that a later prompt would silently rewrite.
 */
function selectedImage(
  drawings: ISlideDrawingService,
  unitId: string,
  snapshot: ISlideDrawingStateSnapshot
): PanelSelection | null {
  const context = snapshot.context
  // A page owned by another Unit is a different Viewer's browser state.
  if (context === null || context.unitId !== unitId) return null
  if (snapshot.selectedIds.length !== 1) return null
  const elementId = snapshot.focusedId ?? snapshot.selectedIds[0]
  if (elementId === null || elementId === undefined) return null
  const drawing = drawings.getDrawingByParam({
    unitId,
    subUnitId: context.subUnitId,
    drawingId: elementId
  })
  if (drawing?.element.type !== PageElementTypeEnum.Image) return null
  return { pageId: context.subUnitId, elementId }
}

/** Post one report; a lost report only means the Host keeps the previous selection. */
async function postSelection(body: unknown): Promise<void> {
  try {
    await fetch('/univer-api/panel/select', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body)
    })
  } catch {
    // The Host may be reloading or the DSH connection may be closed; the next selection re-reports.
  }
}
