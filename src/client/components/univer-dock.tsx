import * as React from 'react'
import { createPortal } from 'react-dom'
import type { SessionListState } from '@deepseek-ai/dsh-api-session-controller/client'
import type { ChatSnapshot } from '@deepseek-ai/dsh-client-ui-chat/client'
import type { ConversationTimelineSnapshot } from '@deepseek-ai/dsh-client-ui-conversation/client'
import type { PropsLocale, PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import { getPanel, openPanel } from '../api/univer-api.ts'
import {
  opensPanel,
  turnFilesOfTimeline,
  type UniverTurnOperation
} from '../conversation/univer-turn-definition.ts'
import { useGatewayStatus, useUniverStates } from '../hooks/use-univer-state.ts'
import { keepsOpenPresentation } from '../open-presentation.ts'
import { addressOfOpenPreview } from '../preview-registry.ts'
import { sessionFileAddress } from '../resource-address.ts'
import type { ViewerLocaleInjected } from '../viewer-locale.ts'
import { UniverPanel } from './univer-panel.tsx'

/**
 * What the input-dock seat hands this plugin.
 *
 * `hasSidebarPreview` reports whether the host can show a file in the right
 * Sidebar, and `openPreview` reveals it by address. Both are absent on a host
 * without that Sidebar, which is the only case that still renders the docked
 * panel.
 */
interface UniverDockInjected extends ViewerLocaleInjected {
  readonly hasSidebarPreview: () => boolean
  readonly openPreview: (address: string) => void
}

interface UniverDockShared extends PropsLocale<'univer'>, UniverDockInjected {}

export type UniverDockProps = PropsRuntime<'conversation.input.dock'> & UniverDockShared

interface PanelTarget {
  readonly file: string
  readonly worktreeId: string | null
  readonly preferredUnitId: string | null
}

/**
 * Route a conversation's Univer work to the one surface the host has.
 *
 * With a right Sidebar, that surface is the preview tab the user opens by
 * clicking the file: an agent write binds the Host target and then reveals that
 * tab, which adopts the binding in place. Nothing is rendered here, so no second
 * panel appears beside the preview and no review copy of the file exists.
 *
 * Without one, the docked panel below the conversation is the only surface left,
 * and it keeps the previous behaviour: the newest write target takes it over and
 * the user reviews there.
 */
export function UniverDock(props: UniverDockProps): React.ReactElement {
  const timeline = props.useChat((snapshot: ChatSnapshot) => snapshot.timeline)
  const cwd = props.useSessions((state: SessionListState) => state.byId[props.sessionId]?.cwd)
  if (props.hasSidebarPreview()) {
    return (
      <UniverPreviewOpener
        key={props.sessionId}
        sessionId={props.sessionId}
        cwd={cwd}
        timeline={timeline}
        openPreview={props.openPreview}
      />
    )
  }
  return <UniverSessionDock key={props.sessionId} {...props} timeline={timeline} cwd={cwd} />
}

/**
 * Reveal the single preview for every write-driven target, without rendering.
 *
 * The Host target is bound first so the preview shows the exact worktree and
 * Unit the write named; a ready worktree, for example, must reopen as `ready`
 * rather than as whatever draft the preview would otherwise resolve. The address
 * comes from the preview that currently shows this container when the user has
 * one open, so a deck opened as `.pptx` is revealed as that same tab instead of
 * gaining a second one for its `.univer` container.
 */
function UniverPreviewOpener(props: {
  readonly sessionId: SessionId
  readonly cwd: string | undefined
  readonly timeline: ConversationTimelineSnapshot | undefined
  readonly openPreview: (address: string) => void
}): React.ReactElement {
  const turnFiles = React.useMemo(
    () => turnFilesOfTimeline(props.timeline, props.cwd),
    [props.timeline, props.cwd]
  )
  const [request, setRequest] = React.useState<PanelTarget | null>(null)
  const seen = React.useRef(new Set<string>())

  // Only operations that deliberately open a surface count; reads never reveal anything.
  React.useEffect(() => {
    let latest: PanelTarget | null = null
    for (const file of turnFiles) {
      for (const operation of file.operations) {
        if (operation.phase === 'failed' || !opensPanel(operation)) continue
        if (seen.current.has(operation.callId)) continue
        seen.current.add(operation.callId)
        const candidate = panelTargetOf(operation, file.file)
        if (candidate !== null) latest = candidate
      }
    }
    if (latest !== null) setRequest(latest)
  }, [turnFiles])

  const { sessionId, cwd, openPreview } = props
  React.useEffect(() => {
    if (request === null) return
    // A presentation tab is already showing the deck the user opened. Writes that name some
    // other file — a regenerated pptx, a second container — must not rebind the session or
    // open another tab. Edits of this deck still fall through and reveal the same address.
    if (keepsOpenPresentation(sessionId, request.file)) return
    void (async () => {
      // The Host answers which file this tab represents: a presentation stays the `.pptx` the user
      // clicked even when the write and the container are named differently. A closed tab therefore
      // reopens as the deck, never as the container that caches it.
      let represented = request.file
      try {
        if (request.worktreeId !== null) {
          const bound = await openPanel(request.file, sessionId, {
            worktreeId: request.worktreeId,
            ...(request.preferredUnitId === null ? {} : { unitId: request.preferredUnitId })
          })
          if (bound.ok && bound.panel.source !== null) represented = bound.panel.source
        } else {
          const bound = await getPanel(sessionId)
          if (bound.panel?.source != null) represented = bound.panel.source
        }
      } catch {
        // The reveal below still opens the preview; it then resolves the file itself and
        // reports a real failure in place instead of hiding it behind this binding attempt.
      }
      openPreview(
        addressOfOpenPreview(sessionId, request.file) ??
          sessionFileAddress(sessionId, cwd, represented)
      )
    })()
  }, [request, sessionId, cwd, openPreview])

  return <></>
}

/** The fallback panel for hosts with no right Sidebar to reveal. */
function UniverSessionDock(
  props: UniverDockShared & {
    readonly sessionId: SessionId
    readonly timeline: ConversationTimelineSnapshot | undefined
    readonly cwd: string | undefined
  }
): React.ReactElement {
  const turnFiles = React.useMemo(
    () => turnFilesOfTimeline(props.timeline, props.cwd),
    [props.timeline, props.cwd]
  )
  const [target, setTarget] = React.useState<PanelTarget | null>(null)
  const seen = React.useRef(new Set<string>())

  // Every write-driven target joins the single panel; only an explicit close removes it.
  React.useEffect(() => {
    let latest: PanelTarget | null = null
    for (const file of turnFiles) {
      for (const operation of file.operations) {
        if (operation.phase === 'failed' || !opensPanel(operation)) continue
        if (seen.current.has(operation.callId)) continue
        seen.current.add(operation.callId)
        const candidate = panelTargetOf(operation, file.file)
        if (candidate !== null) latest = candidate
      }
    }
    if (latest !== null) setTarget(latest)
  }, [turnFiles])

  const { states, applyState } = useUniverStates(
    React.useMemo(() => (target === null ? [] : [target.file]), [target]),
    props.sessionId
  )
  const gateway = useGatewayStatus()

  if (target === null) return <></>
  // The input dock renders inside a translucent, non-draggable container.
  // Portaling avoids inheriting that container's opacity and hit-testing.
  return createPortal(
    <UniverPanel
      file={target.file}
      state={states[target.file]}
      preferredUnitId={target.preferredUnitId}
      gatewayPhase={gateway.phase}
      startGateway={gateway.start}
      sessionId={props.sessionId}
      t={props.t}
      viewerLocale={props.getViewerLocale()}
      onState={applyState}
      onClose={() => setTarget(null)}
    />,
    document.body
  )
}

function panelTargetOf(operation: UniverTurnOperation, file: string): PanelTarget | null {
  if (operation.name === 'new') return { file, worktreeId: null, preferredUnitId: operation.unitId }
  if (operation.worktreeId === null) return null
  return { file, worktreeId: operation.worktreeId, preferredUnitId: operation.unitId }
}
