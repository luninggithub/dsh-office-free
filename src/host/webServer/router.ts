import type { IncomingMessage, ServerResponse } from 'node:http'
import type { SessionStore } from '@deepseek-ai/dsh-session'
import { UniverError } from '../service/errors.ts'
import type { UniverService } from '../service/univer-service.ts'
import type { ConnectionTrust } from './viewer-proxy.ts'
import { gatewayStartRoute } from './routes/gateway.ts'
import {
  panelCloseRoute,
  panelImageEditRoute,
  panelOpenRoute,
  panelSelectRoute,
  panelStateRoute
} from './routes/panel.ts'
import { stateRoute } from './routes/state.ts'
import { statusRoute } from './routes/status.ts'
import { worktreeActionRoute } from './routes/worktree-action.ts'

const MAX_BODY_BYTES = 64 * 1024

/** Create the `/univer-api` HTTP dispatcher behind the DSH connection trust fence. */
export function createUniverRouter(
  service: UniverService,
  sessions: SessionStore,
  connection: ConnectionTrust
) {
  return async (request: IncomingMessage, response: ServerResponse): Promise<void> => {
    const rejection = connection.requestRejection(request)
    if (rejection !== undefined) {
      sendJson(response, rejection, {
        ok: false,
        code: rejection === 403 ? 'FORBIDDEN' : 'UNAUTHORIZED',
        message: 'DSH browser authentication is required'
      })
      return
    }
    try {
      const url = new URL(request.url ?? '/', 'http://localhost')
      if (request.method === 'GET' && url.pathname === '/univer-api/status') {
        sendJson(response, 200, await statusRoute(service))
        return
      }
      if (request.method === 'POST' && url.pathname === '/univer-api/gateway/start') {
        sendJson(response, 200, await gatewayStartRoute(service))
        return
      }
      if (request.method === 'GET' && url.pathname === '/univer-api/state') {
        sendJson(
          response,
          200,
          await stateRoute(
            service,
            sessions,
            url.searchParams.get('file'),
            url.searchParams.get('sessionId')
          )
        )
        return
      }
      if (request.method === 'POST' && url.pathname === '/univer-api/worktree-action') {
        sendJson(
          response,
          200,
          await worktreeActionRoute(service, sessions, await readJsonBody(request))
        )
        return
      }
      if (request.method === 'POST' && url.pathname === '/univer-api/panel/open') {
        sendJson(
          response,
          200,
          await panelOpenRoute(service, sessions, await readJsonBody(request))
        )
        return
      }
      if (request.method === 'POST' && url.pathname === '/univer-api/panel/close') {
        sendJson(response, 200, panelCloseRoute(service, await readJsonBody(request)))
        return
      }
      if (request.method === 'POST' && url.pathname === '/univer-api/panel/select') {
        sendJson(response, 200, panelSelectRoute(service, sessions, await readJsonBody(request)))
        return
      }
      if (request.method === 'POST' && url.pathname === '/univer-api/panel/image-edit') {
        const controller = new AbortController()
        const abortIfDisconnected = (): void => {
          if (!response.writableEnded) controller.abort()
        }
        response.on('close', abortIfDisconnected)
        try {
          sendJson(
            response,
            200,
            await panelImageEditRoute(
              service,
              sessions,
              await readJsonBody(request),
              controller.signal
            )
          )
        } finally {
          response.removeListener('close', abortIfDisconnected)
        }
        return
      }
      if (request.method === 'GET' && url.pathname === '/univer-api/panel') {
        sendJson(response, 200, panelStateRoute(service, url.searchParams.get('sessionId')))
        return
      }
      response.writeHead(404)
      response.end()
    } catch (error) {
      const rejected =
        error instanceof UniverError &&
        (error.code === 'INVALID_REQUEST' ||
          error.code === 'INVALID_FILE_PATH' ||
          error.code === 'FILE_PERMISSION_DENIED' ||
          error.code === 'IMPORT_FORMAT_UNSUPPORTED' ||
          error.code === 'UNIT_NOT_FOUND' ||
          error.code === 'WORKTREE_NOT_FOUND' ||
          error.code === 'SESSION_SCOPE_UNAVAILABLE' ||
          error.code === 'SESSION_SCOPE_DENIED' ||
          error.code === 'IMAGE_SELECTION_MISSING')
      const forbidden =
        error instanceof UniverError &&
        (error.code === 'FILE_PERMISSION_DENIED' || error.code === 'SESSION_SCOPE_DENIED')
      const status = rejected ? (forbidden ? 403 : 400) : 500
      sendJson(response, status, {
        ok: false,
        code: error instanceof UniverError ? error.code : 'INTERNAL_ERROR',
        message: error instanceof Error ? error.message : String(error)
      })
    }
  }
}

/** Send a JSON response with no browser cache. */
export function sendJson(response: ServerResponse, status: number, value: unknown): void {
  response.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store'
  })
  response.end(JSON.stringify(value))
}

async function readJsonBody(request: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = []
  let bytes = 0
  for await (const chunk of request) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    bytes += buffer.length
    if (bytes > MAX_BODY_BYTES)
      throw new UniverError('request body is too large', 'INVALID_REQUEST')
    chunks.push(buffer)
  }
  if (chunks.length === 0) throw new UniverError('JSON body is required', 'INVALID_REQUEST')
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8')) as unknown
  } catch (error) {
    throw new UniverError('request body must be valid JSON', 'INVALID_REQUEST', { cause: error })
  }
}
