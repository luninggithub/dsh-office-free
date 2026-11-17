import * as React from 'react'
import type { TranslateNS } from '@deepseek-ai/dsh-client-ui-slots'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { FileState } from '../../shared/wire/state.ts'
import { closePanel } from '../api/univer-api.ts'
import type { ViewerLocale } from '../viewer-locale.ts'
import { UniverReviewBody, useUniverReview } from './univer-review.tsx'

const WIDTH_STORAGE_KEY = 'dsh-univer-office.panel-width'
const MIN_WIDTH = 360
const DEFAULT_WIDTH = 560

/**
 * The docked Univer panel: the fallback surface for hosts without a right Sidebar.
 *
 * It occupies the browser's right edge at the same page depth as the conversation and resizes from
 * its left splitter. Every Unit kind renders here, and while mounted the panel holds the session's
 * target binding so a later message continues this exact file, worktree, and Unit.
 *
 * A host that provides the right Sidebar never mounts this: there the single Sidebar preview is the
 * only Univer surface, and an agent write reveals that tab instead of opening this panel beside it.
 * The review controls themselves are shared with that preview through `UniverReviewBody`, so the
 * fallback cannot drift from the primary path.
 */
export function UniverPanel(props: {
  readonly file: string
  readonly state: FileState | undefined
  readonly preferredUnitId: string | null
  readonly gatewayPhase: 'checking' | 'stopped' | 'starting' | 'running' | 'failed'
  readonly startGateway: () => Promise<void>
  readonly sessionId: SessionId
  readonly t: TranslateNS<'univer'>
  readonly viewerLocale: ViewerLocale
  readonly onState: (state: FileState) => void
  readonly onClose: () => void
}): React.ReactElement {
  const [width, setWidth] = React.useState(readStoredWidth)
  const [dragging, setDragging] = React.useState(false)
  const [closePrompt, setClosePrompt] = React.useState(false)
  const review = useUniverReview({
    ...props,
    closePrompt,
    dismissClosePrompt: () => setClosePrompt(false),
    closeTab: props.onClose
  })
  // The panel's own close button refuses to drop unsaved work: it becomes the same question the
  // right-Sidebar tab asks before it closes.
  const requestClose = (): void => {
    if (review.unsaved) setClosePrompt(true)
    else props.onClose()
  }

  // Closing the whole panel releases the binding; switching versions or Units inside it never does.
  React.useEffect(
    () => () => {
      void closePanel(props.sessionId).catch(() => {
        // Unload best effort: a dead connection already dropped the Host-side binding.
      })
    },
    [props.sessionId]
  )

  const maxWidth = React.useCallback(
    () => Math.max(MIN_WIDTH, Math.min(Math.round(window.innerWidth * 0.72), 1100)),
    []
  )

  const onPointerDown = (event: React.PointerEvent<HTMLDivElement>): void => {
    event.preventDefault()
    event.currentTarget.setPointerCapture(event.pointerId)
    setDragging(true)
    document.body.style.userSelect = 'none'
  }
  const onPointerMove = (event: React.PointerEvent<HTMLDivElement>): void => {
    if (!dragging) return
    setWidth(clampWidth(window.innerWidth - event.clientX, maxWidth()))
  }
  const onPointerUp = (event: React.PointerEvent<HTMLDivElement>): void => {
    if (event.currentTarget.hasPointerCapture(event.pointerId))
      event.currentTarget.releasePointerCapture(event.pointerId)
    setDragging(false)
    document.body.style.userSelect = ''
    storeWidth(width)
  }

  React.useEffect(
    () => () => {
      document.body.style.userSelect = ''
    },
    []
  )

  return (
    <aside
      className={`uvf_dock${dragging ? ' uvf_dock_dragging' : ''}`}
      style={{ width }}
      data-status={review.unsaved ? 'unsaved' : 'saved'}
      data-file={props.file}
      aria-label={review.title}
    >
      <hr
        className="uvf_dockSplitter"
        aria-orientation="vertical"
        aria-label={props.t('dock.resize')}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
      />
      <header className="uvf_dockHead">
        <span className="uvf_dockGlyph" aria-hidden="true">
          <PanelMark />
        </span>
        <span className="uvf_dockIdentity">
          <span className="uvf_dockTitle" title={review.title}>
            {review.title}
          </span>
          <span className="uvf_dockFile" title={props.file}>
            {props.file}
          </span>
        </span>
        <span className="uvf_dockChip" data-status={review.unsaved ? 'unsaved' : 'saved'}>
          <span className="uvf_dockDot" aria-hidden="true" />
          {review.statusLabel}
        </span>
        <button
          type="button"
          className="uvf_btn"
          title={props.t('dock.refresh')}
          aria-label={props.t('dock.refresh')}
          onClick={review.refresh}
        >
          <RefreshIcon />
        </button>
        <button
          type="button"
          className="uvf_btn"
          title={props.t('dock.close')}
          aria-label={props.t('dock.close')}
          onClick={requestClose}
        >
          <CloseIcon />
        </button>
      </header>
      <UniverReviewBody review={review} t={props.t} />
    </aside>
  )
}

function clampWidth(value: number, max: number): number {
  return Math.min(Math.max(Math.round(value), MIN_WIDTH), max)
}

function readStoredWidth(): number {
  if (typeof window === 'undefined') return DEFAULT_WIDTH
  const stored = Number(window.localStorage.getItem(WIDTH_STORAGE_KEY))
  const max = Math.max(MIN_WIDTH, Math.min(Math.round(window.innerWidth * 0.72), 1100))
  return Number.isFinite(stored) && stored > 0
    ? clampWidth(stored, max)
    : clampWidth(DEFAULT_WIDTH, max)
}

function storeWidth(width: number): void {
  try {
    window.localStorage.setItem(WIDTH_STORAGE_KEY, String(width))
  } catch {
    // Private-mode storage failures must not block resizing.
  }
}

function PanelMark(): React.ReactElement {
  return (
    <svg viewBox="0 0 20 20">
      <rect x="3.5" y="4.5" width="13" height="9" rx="1.6" />
      <path d="M10 13.5v3m-3 0h6M7 7.5h6M7 10.5h4" />
    </svg>
  )
}

function RefreshIcon(): React.ReactElement {
  return (
    <svg viewBox="0 0 16 16" aria-hidden="true">
      <path d="M13 8a5 5 0 1 1-1.5-3.6M13 3v3h-3" />
    </svg>
  )
}

function CloseIcon(): React.ReactElement {
  return (
    <svg viewBox="0 0 16 16" aria-hidden="true">
      <path d="m4 4 8 8M12 4l-8 8" />
    </svg>
  )
}
