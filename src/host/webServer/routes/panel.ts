import { basename, dirname, extname, join } from 'node:path'
import { SessionId, type SessionStore } from '@deepseek-ai/dsh-session'
import type { OpenPanel, PanelOpenResult } from '../../../shared/wire/panel.ts'
import { UniverError } from '../../service/errors.ts'
import {
  unitId as brandedUnitId,
  univerFilePath,
  worktreeId as brandedWorktreeId
} from '../../service/identifiers.ts'
import type { UniverService } from '../../service/univer-service.ts'
import { resolveExistingUniverPath, resolveExistingWorkspacePath } from '../../service/workspace.ts'

/** Body of one panel open, close, or selection request from the browser. */
interface PanelBody {
  readonly file?: unknown
  readonly sessionId?: unknown
  readonly unitId?: unknown
  readonly worktreeId?: unknown
  readonly pageId?: unknown
  readonly elementId?: unknown
}

/**
 * Open the file the user clicked in the Univer side panel.
 *
 * A `.pptx` is imported once into a sibling `.univer` container, and a `.univer` is bound to its
 * existing Unit, so either click ends in the same editable target. The Host authorizes the source
 * against the named session's workspace before any Gateway write.
 */
export async function panelOpenRoute(
  service: UniverService,
  sessions: SessionStore,
  body: unknown
): Promise<PanelOpenResult> {
  const { file, sessionId, unitId, worktreeId } = requireBody(body)
  const cwd = sessionWorkspace(sessionId, sessions)
  const extension = extname(file).toLowerCase()
  if (extension === '.pptx') {
    const source = await resolveExistingWorkspacePath(cwd, file)
    const target = join(
      dirname(source.path),
      `${basename(source.path, extname(source.path))}.univer`
    )
    const result = await service.openPanel({
      workspace: source.workspace,
      file: univerFilePath(target),
      source: source.path,
      sourceWorkspace: source.workspace,
      sessionId
    })
    return scopeResult(result, sessionId)
  }
  if (extension === '.univer') {
    const target = await resolveExistingUniverPath(cwd, file)
    const result = await service.bindPanel({
      workspace: target.workspace,
      file: target.path,
      sessionId,
      ...(worktreeId === undefined ? {} : { worktreeId: brandedWorktreeId(worktreeId) }),
      ...(unitId === undefined ? {} : { unitId: brandedUnitId(unitId) })
    })
    return scopeResult(result, sessionId)
  }
  throw new UniverError('The Univer panel opens .pptx and .univer files only.', 'INVALID_REQUEST')
}

/** Release the session's panel focus when its panel closes. */
export function panelCloseRoute(service: UniverService, body: unknown): { readonly ok: true } {
  const { sessionId } = requireSession(body)
  service.releasePanel(sessionId)
  return { ok: true }
}

/**
 * Record which image the Viewer has selected, so "edit this image" can name it.
 *
 * Only an already-bound session keeps the report: the panel's scope is the session workspace, so an
 * unknown session has no Viewer and nothing can be selected there.
 */
export function panelSelectRoute(
  service: UniverService,
  sessions: SessionStore,
  body: unknown
): { readonly ok: true } {
  const { sessionId, pageId, elementId } = requireSelection(body)
  sessionWorkspace(sessionId, sessions)
  service.reportPanelSelection({
    sessionId,
    ...(pageId === undefined ? {} : { pageId }),
    ...(elementId === undefined ? {} : { elementId })
  })
  return { ok: true }
}

/** Read the panel target a session currently holds, for panel remounts. */
export function panelStateRoute(
  service: UniverService,
  sessionId: unknown
): { readonly ok: true; readonly panel: OpenPanel | null } {
  const id = requireSessionId(sessionId)
  const panel = service.panelOf(id)
  return {
    ok: true,
    panel: panel === undefined ? null : scopePanel(panel, id)
  }
}

function requireBody(body: unknown): {
  readonly file: string
  readonly sessionId: string
  readonly unitId: string | undefined
  readonly worktreeId: string | undefined
} {
  if (typeof body !== 'object' || body === null || Array.isArray(body))
    throw new UniverError('JSON object body is required', 'INVALID_REQUEST')
  const value = body as PanelBody
  const sessionId = requireSessionId(value.sessionId)
  if (typeof value.file !== 'string' || value.file.length === 0)
    throw new UniverError('file is required', 'INVALID_REQUEST')
  return {
    file: value.file,
    sessionId,
    unitId: typeof value.unitId === 'string' && value.unitId.length > 0 ? value.unitId : undefined,
    worktreeId:
      typeof value.worktreeId === 'string' && value.worktreeId.length > 0
        ? value.worktreeId
        : undefined
  }
}

function requireSession(body: unknown): { readonly sessionId: string } {
  if (typeof body !== 'object' || body === null || Array.isArray(body))
    throw new UniverError('JSON object body is required', 'INVALID_REQUEST')
  return { sessionId: requireSessionId((body as PanelBody).sessionId) }
}

/**
 * Read one selection report.
 *
 * Both ids together name an element; either one alone leaves the report incomplete, which is the
 * same browser state as "nothing selected" for the only consumer.
 */
function requireSelection(body: unknown): {
  readonly sessionId: string
  readonly pageId: string | undefined
  readonly elementId: string | undefined
} {
  if (typeof body !== 'object' || body === null || Array.isArray(body))
    throw new UniverError('JSON object body is required', 'INVALID_REQUEST')
  const value = body as PanelBody
  const sessionId = requireSessionId(value.sessionId)
  const pageId = readId(value.pageId)
  const elementId = readId(value.elementId)
  if (pageId === undefined || elementId === undefined)
    return { sessionId, pageId: undefined, elementId: undefined }
  return { sessionId, pageId, elementId }
}

function readId(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined
}

function requireSessionId(value: unknown): string {
  if (typeof value !== 'string' || value.length === 0)
    throw new UniverError('sessionId is required', 'INVALID_REQUEST')
  return value
}

function sessionWorkspace(sessionId: string, sessions: SessionStore): string {
  const cwd = sessions.get(SessionId(sessionId))?.header.cwd
  if (cwd === undefined)
    throw new UniverError('session is unavailable or has no workspace', 'SESSION_SCOPE_UNAVAILABLE')
  return cwd
}

function scopeResult(result: PanelOpenResult, sessionId: string): PanelOpenResult {
  return result.ok ? { ok: true, panel: scopePanel(result.panel, sessionId) } : result
}

/** Viewer document opens are authorized against the named session, so every URL carries it. */
function scopePanel(panel: OpenPanel, sessionId: string): OpenPanel {
  return {
    ...panel,
    viewerUrl: `${panel.viewerUrl}&sessionId=${encodeURIComponent(sessionId)}`
  }
}
