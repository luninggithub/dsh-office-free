import * as React from 'react'
import type { TranslateNS } from '@deepseek-ai/dsh-client-ui-slots'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type { BranchUnit, FileState } from '../../shared/wire/state.ts'
import { openPanel, worktreeAction } from '../api/univer-api.ts'
import { basename } from '../conversation/univer-turn-definition.ts'
import { localizeViewerUrl, type ViewerLocale } from '../viewer-locale.ts'
import { UnitChips, unitViewerUrl } from './unit-chips.tsx'

/** Everything the review controls read, in one object. */
export interface ReviewInputs {
  /** Absolute `.univer` container that owns the displayed Unit. */
  readonly file: string
  readonly state: FileState | undefined
  /** The Unit a conversation write named, until the user picks another. */
  readonly preferredUnitId: string | null
  readonly gatewayPhase: 'checking' | 'stopped' | 'starting' | 'running' | 'failed'
  readonly startGateway: () => Promise<void>
  readonly sessionId: SessionId
  readonly t: TranslateNS<'univer'>
  readonly viewerLocale: ViewerLocale
  readonly onState: (state: FileState) => void
  /**
   * A tab close found unsaved changes. The close was refused, and the surface now shows the
   * confirm strip, so the user decides instead of losing work.
   */
  readonly closePrompt: boolean
  /** Dismiss the close-confirm strip without closing. */
  readonly dismissClosePrompt: () => void
  /** Close the tab once the close-confirm strip is settled. */
  readonly closeTab: () => void
}

/** The branch's unsaved state, and every action its toolbar can take. */
export interface Review {
  readonly file: string
  readonly title: string
  readonly statusLabel: string
  /** True while the branch differs from the last saved version. */
  readonly unsaved: boolean
  readonly units: readonly BranchUnit[]
  readonly unitId: string | undefined
  readonly editing: boolean
  readonly compare: boolean
  readonly pending: 'save' | 'discard' | null
  readonly confirmingDiscard: boolean
  readonly closePrompt: boolean
  readonly failure: string | null
  /** The Viewer target for the branch, Unit, and comparison. */
  readonly url: string | undefined
  /** Changes whenever the Viewer must be reloaded in place. */
  readonly reloadKey: number
  readonly gatewayPhase: ReviewInputs['gatewayPhase']
  selectUnit: (unitId: string | undefined) => void
  toggleCompare: () => void
  save: () => Promise<void>
  requestDiscard: () => void
  confirmDiscard: () => Promise<void>
  cancelDiscard: () => void
  resolveClose: (decision: 'save' | 'discard' | 'cancel') => Promise<void>
  refresh: () => void
  startGateway: () => Promise<void>
}

/**
 * Resolve one file, its single branch, and a Unit into review state.
 *
 * The model is deliberately small: a file has a last saved version and at most one branch of
 * unsaved changes. Editing — by the user in this Viewer or by an agent through the Host — lands in
 * that branch, which is what makes Save and Discard the only two decisions the toolbar offers.
 * Neither is automatic: Save folds the branch into the saved version, Discard throws it away.
 *
 * The controls live here rather than in the Viewer's own chrome because the Viewer runs with
 * `mode=embedded` and has no review UI of its own. Both surfaces that show it — the right Sidebar
 * preview and the fallback dock — render the same state, so the two cannot diverge.
 */
export function useUniverReview(inputs: ReviewInputs): Review {
  const [nonce, setNonce] = React.useState(0)
  const [selected, setSelected] = React.useState<string | undefined>(
    inputs.preferredUnitId ?? undefined
  )
  const [compare, setCompare] = React.useState(false)
  const [pending, setPending] = React.useState<'save' | 'discard' | null>(null)
  const [confirmingDiscard, setConfirmingDiscard] = React.useState(false)
  const [failure, setFailure] = React.useState<string | null>(null)
  const bindKey = React.useRef<string | null>(null)
  const branch = inputs.state?.branch ?? null

  // A new branch replaces what the surface shows; a user's own Unit choice survives every poll, so
  // only a changed branch identity resets it.
  const signature = `${branch?.worktreeId ?? ''}\u0000${inputs.preferredUnitId ?? ''}`
  const lastSignature = React.useRef(signature)
  React.useEffect(() => {
    if (lastSignature.current === signature) return
    lastSignature.current = signature
    setSelected(inputs.preferredUnitId ?? undefined)
    setCompare(false)
    setConfirmingDiscard(false)
    setFailure(null)
    setNonce((value) => value + 1)
  }, [signature, inputs.preferredUnitId])

  const units = branch?.units ?? []
  const unitId =
    selected !== undefined && units.some((unit) => unit.unitId === selected)
      ? selected
      : inputs.preferredUnitId !== null &&
          units.some((unit) => unit.unitId === inputs.preferredUnitId)
        ? inputs.preferredUnitId
        : units[0]?.unitId
  const unit = units.find((entry) => entry.unitId === unitId)
  const unsaved = branch?.unsaved === true

  // The Viewer target follows the branch, the Unit, and the diff toggle. A file with no branch is
  // showing the saved version, which is read-only here: without a branch there is nothing to undo.
  const url = React.useMemo(() => {
    const target =
      branch === null
        ? trunkTarget(inputs.state)
        : unitViewerUrl(branch.worktreeUrl, branch.units, unitId)
    if (target === undefined) return undefined
    const next = new URL(target, window.location.origin)
    if (branch !== null && compare && unsaved) next.searchParams.set('view', 'diff')
    return localizeViewerUrl(next.toString(), inputs.viewerLocale)
  }, [branch, compare, inputs.state, inputs.viewerLocale, unitId, unsaved])

  // Leaving the page without saving is allowed, but the browser asks first: the branch is not
  // written back on unload, so the next open starts from the last saved version.
  React.useEffect(() => {
    if (!unsaved) return undefined
    const warn = (event: BeforeUnloadEvent): void => {
      event.preventDefault()
      event.returnValue = ''
    }
    window.addEventListener('beforeunload', warn)
    return () => window.removeEventListener('beforeunload', warn)
  }, [unsaved])

  // Hold this branch and Unit as the session's target; the Host projects it through `univer_status`.
  React.useEffect(() => {
    if (branch === null || unitId === undefined) return
    const key = `${inputs.file}\u0000${branch.worktreeId}\u0000${unitId}`
    if (bindKey.current === key) return
    bindKey.current = key
    void openPanel(inputs.file, inputs.sessionId, { worktreeId: branch.worktreeId, unitId }).catch(
      () => {
        // A closed surface or vanished file must not break the conversation; the next poll retries.
        bindKey.current = null
      }
    )
  }, [branch, inputs.file, inputs.sessionId, unitId])

  const run = React.useCallback(
    async (action: 'merge' | 'discard', close: boolean): Promise<boolean> => {
      const target = inputs.state?.branch ?? null
      if (target === null || pending !== null) return false
      setPending(action === 'merge' ? 'save' : 'discard')
      setFailure(null)
      setConfirmingDiscard(false)
      try {
        const result = await worktreeAction(
          inputs.file,
          inputs.sessionId,
          target.worktreeId,
          action
        )
        if (!result.ok) {
          setFailure(result.reason)
          if (result.state !== undefined) inputs.onState(result.state)
          return false
        }
        inputs.onState(result.state)
        setCompare(false)
        // Saving or discarding settles this branch; unless the tab is closing, the Host opens the
        // next one so the user can keep editing immediately.
        if (!close) await nextBranch(inputs).catch(() => undefined)
        setNonce((value) => value + 1)
        return true
      } catch (error) {
        setFailure(error instanceof Error ? error.message : String(error))
        return false
      } finally {
        setPending(null)
      }
    },
    [inputs, pending]
  )

  return {
    file: inputs.file,
    title: unit?.name || branch?.name || basename(inputs.file),
    statusLabel: unsaved ? inputs.t('dock.unsaved') : inputs.t('dock.saved'),
    unsaved,
    units,
    unitId,
    editing: branch !== null,
    compare,
    pending,
    confirmingDiscard,
    closePrompt: inputs.closePrompt,
    failure,
    url,
    reloadKey: nonce,
    gatewayPhase: inputs.gatewayPhase,
    selectUnit: setSelected,
    toggleCompare: () => setCompare((value) => !value),
    save: async () => {
      await run('merge', false)
    },
    requestDiscard: () => setConfirmingDiscard(true),
    confirmDiscard: async () => {
      await run('discard', false)
    },
    cancelDiscard: () => setConfirmingDiscard(false),
    resolveClose: async (decision) => {
      if (decision === 'cancel') {
        inputs.dismissClosePrompt()
        return
      }
      const settled = await run(decision === 'save' ? 'merge' : 'discard', true)
      if (!settled) return
      inputs.dismissClosePrompt()
      inputs.closeTab()
    },
    refresh: () => setNonce((value) => value + 1),
    startGateway: inputs.startGateway
  }
}

/** Bind the branch that holds the next round of edits, once the previous one is settled. */
async function nextBranch(inputs: ReviewInputs): Promise<void> {
  const result = await openPanel(inputs.file, inputs.sessionId)
  if (!result.ok) throw new Error(result.reason)
}

/**
 * The review controls and the embedded Viewer.
 *
 * This is the body both surfaces share: the dock adds its own frame and header around it, and the
 * right-Sidebar preview renders it directly under the tab's chrome. Every class,
 * `data-panel-action`, and note is identical, so a control behaves the same wherever it is shown.
 */
export function UniverReviewBody(props: {
  readonly review: Review
  readonly t: TranslateNS<'univer'>
  /** Show the saved/unsaved chip inline; the dock already carries one in its header. */
  readonly showStatus?: boolean
}): React.ReactElement {
  const { review, t } = props
  const settled = review.pending !== null
  return (
    <>
      <div className="uvf_dockBar" role="toolbar" aria-label={t('dock.actions')}>
        {props.showStatus === true ? (
          <span className="uvf_dockChip uvf_surfaceChip" data-status={review.unsaved ? 'draft' : 'saved'}>
            <span className="uvf_dockDot" aria-hidden="true" />
            {review.statusLabel}
          </span>
        ) : null}
        {review.unsaved ? (
          <button
            type="button"
            className={`uvf_dockAction${review.compare ? ' uvf_dockAction_on' : ''}`}
            data-panel-action="compare"
            aria-pressed={review.compare}
            disabled={settled}
            onClick={review.toggleCompare}
          >
            {t('dock.compare')}
          </button>
        ) : null}
        <button
          type="button"
          className="uvf_dockAction uvf_dockAction_primary"
          data-panel-action="save"
          disabled={!review.unsaved || settled}
          onClick={() => void review.save()}
        >
          {review.pending === 'save' ? t('dock.saving') : t('dock.save')}
        </button>
        <button
          type="button"
          className="uvf_dockAction uvf_dockAction_danger"
          data-panel-action="discard"
          disabled={!review.unsaved || settled}
          onClick={review.requestDiscard}
        >
          {review.pending === 'discard' ? t('dock.discarding') : t('dock.discard')}
        </button>
        {review.editing ? (
          <span
            className="uvf_dockChip uvf_dockEdit"
            data-status="editable"
            title={t('dock.editableHint')}
          >
            <span className="uvf_dockDot" aria-hidden="true" />
            {t('dock.editable')}
          </span>
        ) : null}
      </div>
      {review.closePrompt ? (
        <div className="uvf_dockConfirm" role="alertdialog" aria-label={t('dock.unsavedClose')}>
          <span>{t('dock.unsavedClose')}</span>
          <button
            type="button"
            className="uvf_dockAction uvf_dockAction_primary"
            data-panel-action="close-save"
            disabled={settled}
            onClick={() => void review.resolveClose('save')}
          >
            {t('dock.saveAndClose')}
          </button>
          <button
            type="button"
            className="uvf_dockAction uvf_dockAction_danger"
            data-panel-action="close-discard"
            disabled={settled}
            onClick={() => void review.resolveClose('discard')}
          >
            {t('dock.discardAndClose')}
          </button>
          <button
            type="button"
            className="uvf_dockAction"
            data-panel-action="close-cancel"
            disabled={settled}
            onClick={() => void review.resolveClose('cancel')}
          >
            {t('dock.cancel')}
          </button>
        </div>
      ) : null}
      {review.confirmingDiscard ? (
        <div className="uvf_dockConfirm" role="alertdialog" aria-label={t('dock.confirmDiscard')}>
          <span>{t('dock.confirmDiscard')}</span>
          <button
            type="button"
            className="uvf_dockAction uvf_dockAction_danger"
            data-panel-action="confirm-discard"
            disabled={settled}
            onClick={() => void review.confirmDiscard()}
          >
            {t('dock.confirm')}
          </button>
          <button
            type="button"
            className="uvf_dockAction"
            data-panel-action="cancel"
            onClick={review.cancelDiscard}
          >
            {t('dock.cancel')}
          </button>
        </div>
      ) : null}
      {review.failure === null ? null : (
        <output className="uvf_dockFailure" data-panel-error="">
          {t('dock.actionFailed')}: {review.failure}
        </output>
      )}
      <div className="uvf_dockBody">
        <UnitChips units={review.units} selected={review.unitId} t={t} onSelect={review.selectUnit} />
        {review.url === undefined ? (
          <div className="uvf_note" data-panel-note={review.gatewayPhase}>
            <span>
              {review.gatewayPhase === 'stopped' || review.gatewayPhase === 'failed'
                ? t('dock.gatewayDown')
                : t('dock.loading')}
            </span>
            {review.gatewayPhase === 'stopped' || review.gatewayPhase === 'failed' ? (
              <button type="button" onClick={() => void review.startGateway()}>
                {t('dock.startGateway')}
              </button>
            ) : null}
          </div>
        ) : (
          <iframe
            key={`${review.url}#${String(review.reloadKey)}`}
            className="uvf_dockFrame"
            src={review.url}
            title={review.title}
          />
        )}
      </div>
    </>
  )
}

/** The saved version, still embedded and still scoped to this browser session. */
function trunkTarget(state: FileState | undefined): string | undefined {
  if (state === undefined || state.viewerUrl === null) return undefined
  const target = new URL(state.viewerUrl, window.location.origin)
  target.searchParams.set('mode', 'embedded')
  target.searchParams.set('scope', 'trunk')
  return target.toString()
}
