import type { SessionId } from '@deepseek-ai/dsh-session/types'
import type {
  WorktreeActionResult,
  WorktreeReviewAction
} from '../../shared/wire/actions.ts'
import type { OpenPanel, PanelOpenResult } from '../../shared/wire/panel.ts'
import type { FileState } from '../../shared/wire/state.ts'
import type { EnsureGatewayResult, UniverStatus } from '../../shared/wire/status.ts'

/** Error envelope returned by the Host browser API. */
interface ApiError {
  readonly message?: string
  readonly code?: string
}

/** Structured Host failure retained for UI decisions that depend on the error code. */
export class UniverApiError extends Error {
  readonly code: string | undefined
  readonly status: number

  constructor(message: string, code: string | undefined, status: number) {
    super(message)
    this.name = 'UniverApiError'
    this.code = code
    this.status = status
  }
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${window.location.origin}${path}`, init)
  const body = (await response.json()) as T | ApiError
  if (!response.ok) {
    const error = body as ApiError
    throw new UniverApiError(
      error.message ?? `Univer API HTTP ${String(response.status)}`,
      error.code,
      response.status
    )
  }
  return body as T
}

/** Read package, Gateway, and Unit content availability. */
export function getUniverStatus(): Promise<UniverStatus> {
  return request('/univer-api/status')
}

/** Start or reuse the bundled Gateway. */
export function startGateway(): Promise<EnsureGatewayResult> {
  return request('/univer-api/gateway/start', { method: 'POST' })
}

/** Read one file's current collaboration state and Viewer targets. */
export function getFileState(file: string, sessionId: SessionId): Promise<FileState> {
  return request(
    `/univer-api/state?file=${encodeURIComponent(file)}&sessionId=${encodeURIComponent(sessionId)}`
  )
}

/**
 * Open the file the user clicked in the Univer panel.
 *
 * A `.pptx` is imported once into its sibling `.univer`; a `.univer` is bound to an existing Unit.
 * Explicit identities let the turn-driven panel bind without re-importing anything.
 */
export function openPanel(
  file: string,
  sessionId: SessionId,
  target: { readonly worktreeId?: string; readonly unitId?: string } = {}
): Promise<PanelOpenResult> {
  return request('/univer-api/panel/open', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      file,
      sessionId,
      ...(target.worktreeId === undefined ? {} : { worktreeId: target.worktreeId }),
      ...(target.unitId === undefined ? {} : { unitId: target.unitId })
    })
  })
}

/** Release the session's panel focus when its panel closes. */
export function closePanel(sessionId: SessionId): Promise<{ ok: true }> {
  return request('/univer-api/panel/close', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ sessionId })
  })
}

/** Read the target a session's right panel currently holds, for panel remounts. */
export function getPanel(sessionId: SessionId): Promise<{ ok: true; panel: OpenPanel | null }> {
  return request(`/univer-api/panel?sessionId=${encodeURIComponent(sessionId)}`)
}

/**
 * Apply one worktree lifecycle decision from the panel header.
 *
 * These are explicit user actions: the Host re-authorizes the file against the session workspace,
 * and the response carries the resulting state so the panel can update without waiting for a poll.
 */
export function worktreeAction(
  file: string,
  sessionId: SessionId,
  worktreeId: string,
  action: WorktreeReviewAction
): Promise<WorktreeActionResult> {
  return request('/univer-api/worktree-action', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ file, sessionId, worktreeId, action })
  })
}

/** A projected file was removed (or never successfully created) in the session workspace. */
export function isMissingUniverFile(error: unknown): boolean {
  return error instanceof UniverApiError && error.code === 'INVALID_FILE_PATH'
}
