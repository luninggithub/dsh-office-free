import * as React from 'react'
import type { Context as ClientContext } from '@deepseek-ai/cordis'
import type { TranslateNS } from '@deepseek-ai/dsh-client-ui-slots'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import { closePanel, getPanel, openPanel } from '../api/univer-api.ts'
import { sameFile } from '../file-key.ts'
import { useGatewayStatus, useUniverStates } from '../hooks/use-univer-state.ts'
import { useUniverLocale } from '../hooks/use-univer-locale.ts'
import { isPresentationPath } from '../presentation-path.ts'
import {
  registerUnsavedGuard,
  type UnsavedPreviewRegistration
} from '../preview-close.ts'
import { registerOpenPreview } from '../preview-registry.ts'
import { UniverReviewBody, useUniverReview } from './univer-review.tsx'

/**
 * How one opened file resolves to a renderable target.
 *
 * `open` carries the container plus the version and Unit held for this session;
 * `trunk` is a `.univer` with no draft to review, which is a mainline preview and
 * not a failure; `failed` is a real error the user can retry.
 */
type Resolution =
  | { readonly kind: 'pending' }
  | {
      readonly kind: 'open'
      readonly file: string
      readonly worktreeId: string | null
      readonly unitId: string | null
    }
  | { readonly kind: 'trunk'; readonly file: string }
  | { readonly kind: 'failed'; readonly reason: string }

interface SessionBinding {
  readonly worktreeId: string | null
  readonly unitId: string | null
}

/**
 * The single Univer preview: every entry point that opens a file renders this.
 *
 * It is the only Univer surface the user sees while the right Sidebar exists —
 * the deck opened from the chat or the file tree, the changes an agent makes to
 * it, and the controls that save or discard those changes all live in this one
 * tab. Nothing opens a second preview beside it.
 *
 * Opening resolves once: a `.pptx` is imported into its sibling `.univer`, and
 * a `.univer` binds to its one editable branch (created from the saved version
 * when the file has none). From then on this surface follows the Host state,
 * which reports that branch, whether it holds unsaved changes, and the Viewer
 * target for its Units.
 */
export function UniverPreviewSurface(props: {
  readonly ctx: ClientContext
  readonly path: string
  readonly sessionId: string
  readonly title: string
  /** The entry point that opened this preview, for diagnostics. */
  readonly surface: string
  /** The address this tab was opened under; it is also the tab's identity. */
  readonly address: string
  /** Close this tab after the user settles unsaved changes; absent on hosts without a close API. */
  readonly closeTab?: (() => void) | undefined
}): React.ReactElement {
  const { t, viewerLocale } = useUniverLocale(props.ctx)
  // Branding the wire session id is this plugin's client-side boundary conversion.
  const sessionId = props.sessionId as SessionId
  const presentation = isPresentationPath(props.path)
  const [resolution, setResolution] = React.useState<Resolution>({ kind: 'pending' })
  const [attempt, setAttempt] = React.useState(0)

  React.useEffect(() => {
    let active = true
    setResolution({ kind: 'pending' })
    // The Host names the container after the deck, so a `.pptx` binding can be matched before the
    // import runs. Adopting it is what keeps a reopened deck on the worktree the agent is editing
    // instead of creating another import draft.
    const containerHint = presentation ? props.path.replace(/\.pptx$/i, '.univer') : props.path
    const settle = (next: Resolution): void => {
      if (active) setResolution(next)
    }
    void (async () => {
      try {
        const bound = (await getPanel(sessionId)).panel
        if (!active) return
        if (bound !== null && bound.worktreeId !== '' && sameFile(bound.file, containerHint)) {
          settle({
            kind: 'open',
            file: bound.file,
            worktreeId: bound.worktreeId,
            unitId: bound.unitId === '' ? null : bound.unitId
          })
          return
        }
        const result = await openPanel(props.path, sessionId)
        if (!active) return
        if (!result.ok) {
          settle({ kind: 'failed', reason: result.reason })
          return
        }
        settle({
          kind: 'open',
          file: result.panel.file,
          worktreeId: result.panel.worktreeId === '' ? null : result.panel.worktreeId,
          unitId: result.panel.unitId === '' ? null : result.panel.unitId
        })
      } catch (error) {
        if (!active) return
        if (presentation) {
          settle({ kind: 'failed', reason: messageOf(error) })
          return
        }
        // A container without a draft is the mainline, not a failure; the state poll decides
        // whether it exists, and the missing/gateway notes cover those cases.
        settle({ kind: 'trunk', file: props.path })
      }
    })()
    return () => {
      active = false
    }
  }, [props.path, props.sessionId, attempt])

  const file =
    resolution.kind === 'open' || resolution.kind === 'trunk' ? resolution.file : undefined
  const { states, missingFiles, applyState } = useUniverStates(
    file === undefined ? [] : [file],
    sessionId
  )
  const gateway = useGatewayStatus()
  const binding = useSessionBinding(file, sessionId)
  // The Host reporting no such file wins over a stale target: the container may have been deleted
  // between the open and the poll, and an empty Viewer would hide that.
  const missing = file !== undefined && missingFiles.has(file)

  // Two panes can show two files; the registry maps every one of them back to its own address.
  React.useEffect(
    () =>
      file === undefined
        ? undefined
        : registerOpenPreview({
            sessionId: props.sessionId,
            container: file,
            address: props.address
          }),
    [file, props.address, props.sessionId]
  )

  // The preview owns the binding while it is mounted: closing the tab hands the conversation back.
  React.useEffect(
    () => () => {
      void closePanel(sessionId).catch(() => {
        // Unload best effort: a dead connection already dropped the Host-side binding.
      })
    },
    [sessionId]
  )

  const initialUnitId = resolution.kind === 'open' ? resolution.unitId : null
  const [closePrompt, setClosePrompt] = React.useState(false)
  const unsaved = states[file ?? '']?.branch?.unsaved === true
  // The close handler runs outside React, so it reads the dirty flag through a ref that every
  // render refreshes; a guard registration survives across polls of the same tab.
  const unsavedRef = React.useRef(unsaved)
  React.useEffect(() => {
    unsavedRef.current = unsaved
  }, [unsaved])
  const registration = React.useRef<UnsavedPreviewRegistration | undefined>(undefined)
  React.useEffect(() => {
    if (file === undefined) return undefined
    const handle = registerUnsavedGuard(props.sessionId, props.address, {
      hasUnsavedChanges: () => unsavedRef.current,
      requestClose: () => setClosePrompt(true)
    })
    registration.current = handle
    return () => {
      registration.current = undefined
      handle.dispose()
    }
  }, [file, props.address, props.sessionId])

  // Settling the branch clears the guard first: the dirty flag lags one poll, and a close issued
  // right after Save must not be refused again.
  const closeTabAction = props.closeTab
  const closeTab = React.useCallback(() => {
    registration.current?.settle()
    closeTabAction?.()
  }, [closeTabAction])

  const review = useUniverReview({
    file: file ?? '',
    state: file === undefined ? undefined : states[file],
    preferredUnitId: binding === undefined ? initialUnitId : binding.unitId,
    gatewayPhase: gateway.phase,
    startGateway: gateway.start,
    sessionId,
    t,
    viewerLocale,
    onState: applyState,
    closePrompt,
    dismissClosePrompt: () => setClosePrompt(false),
    closeTab
  })

  return (
    <div
      className="uvf_fileView"
      data-surface={props.surface}
      data-file={file ?? props.path}
      data-preview-state={file === undefined ? resolution.kind : missing ? 'missing' : 'open'}
    >
      {file === undefined || missing ? (
        <div className="uvf_viewerShell">
          <div className="uvf_note" data-preview-note={missing ? 'missing' : resolution.kind}>
            <span>{noteLabel(resolution, missing, presentation, gateway.phase, t)}</span>
            {resolution.kind === 'failed' && !missing ? (
              <code className="uvf_noteDetail">{resolution.reason}</code>
            ) : null}
            {resolution.kind === 'failed' && !missing ? (
              <button type="button" onClick={() => setAttempt((value) => value + 1)}>
                {t('viewer.retry')}
              </button>
            ) : null}
            {gateway.phase === 'stopped' || gateway.phase === 'failed' ? (
              <button type="button" onClick={() => void gateway.start()}>
                {t('dock.startGateway')}
              </button>
            ) : null}
          </div>
        </div>
      ) : (
        <UniverReviewBody review={review} t={t} showStatus />
      )}
    </div>
  )
}

/**
 * Follow the session's target binding while the surface is mounted.
 *
 * The binding is how the Host reports which file, worktree, and Unit this
 * session is editing, and it is also how a write-driven reveal reaches an
 * already-open preview: the reveal rebinds first, then focuses the tab. A failed
 * poll keeps the last target, because the surface itself still works without it.
 */
function useSessionBinding(
  file: string | undefined,
  sessionId: SessionId,
  intervalMs = 900
): SessionBinding | undefined {
  const [binding, setBinding] = React.useState<SessionBinding | undefined>(undefined)
  React.useEffect(() => {
    if (file === undefined) {
      setBinding(undefined)
      return
    }
    let active = true
    const poll = async (): Promise<void> => {
      try {
        const result = await getPanel(sessionId)
        if (!active) return
        const panel = result.panel
        if (panel === null || panel.worktreeId === '' || !sameFile(panel.file, file)) return
        const next: SessionBinding = {
          worktreeId: panel.worktreeId,
          unitId: panel.unitId === '' ? null : panel.unitId
        }
        setBinding((previous) =>
          previous !== undefined &&
          previous.worktreeId === next.worktreeId &&
          previous.unitId === next.unitId
            ? previous
            : next
        )
      } catch {
        // The binding route is advisory; a transient failure must not blank the preview.
      }
    }
    void poll()
    const timer = window.setInterval(() => void poll(), intervalMs)
    return () => {
      active = false
      window.clearInterval(timer)
    }
  }, [file, sessionId, intervalMs])
  return binding
}

function noteLabel(
  resolution: Resolution,
  missing: boolean,
  presentation: boolean,
  phase: 'checking' | 'stopped' | 'starting' | 'running' | 'failed',
  t: TranslateNS<'univer'>
): string {
  if (missing) return t('viewer.missing')
  if (resolution.kind === 'failed')
    return presentation ? t('viewer.presentationFailed') : t('viewer.missing')
  if (phase === 'stopped' || phase === 'failed') return t('dock.gatewayDown')
  return presentation ? t('viewer.presentationOpening') : t('dock.loading')
}

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
