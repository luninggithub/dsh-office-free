import type { SessionStore } from '@deepseek-ai/dsh-session'
import type { FileState } from '../../../shared/wire/state.ts'
import type { UniverService } from '../../service/univer-service.ts'
import { resolveAuthorizedFile } from '../session-scope.ts'

/** Read one file's current worktree state, with Viewer targets scoped to this browser session. */
export async function stateRoute(
  service: UniverService,
  sessions: SessionStore,
  file: unknown,
  sessionId: unknown
) {
  const authorized = await resolveAuthorizedFile(file, sessionId, sessions)
  const state = await service.fileState({ workspace: authorized.workspace, file: authorized.path })
  // Viewer document opens are authorized against the named session, so the Host must place
  // that session in every projected URL; subsequent Gateway requests use DSH authentication.
  return typeof sessionId === 'string' ? withSessionScope(state, sessionId) : state
}

function withSessionScope(state: FileState, sessionId: string): FileState {
  const query = `sessionId=${encodeURIComponent(sessionId)}`
  const append = (url: string): string => `${url}&${query}`
  const branch = state.branch
  return {
    ...state,
    viewerUrl: state.viewerUrl === null ? null : `${state.viewerUrl}&${query}`,
    branch:
      branch === null
        ? null
        : {
            ...branch,
            worktreeUrl: append(branch.worktreeUrl),
            units: branch.units.map((unit) => ({ ...unit, worktreeUrl: append(unit.worktreeUrl) }))
          }
  }
}
