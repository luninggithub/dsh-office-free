import { createReadStream, existsSync, statSync } from 'node:fs'
import { extname, resolve, sep } from 'node:path'
import { pipeline } from 'node:stream/promises'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('./packages/pptx-kernel/web/', import.meta.url))
const types = new Map([
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.css', 'text/css; charset=utf-8'],
  ['.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
])

export const name = 'dsh-univer-office'

export const inject = ['webServer']

/** Serve the PPTX editor under `/office/` for an authenticated DSH session. */
export function apply(ctx) {
  ctx.effect(() => ctx.webServer.register({
    kind: 'prefix',
    path: '/office',
    auth: 'required',
    handler: (request, response) => serve(request, response),
  }))
  process.stderr.write('dsh-univer-office: PPTX editor at /office/\n')
}

async function serve(request, response) {
  const url = new URL(request.url ?? '/', 'http://127.0.0.1')
  if (request.method !== 'GET' && request.method !== 'HEAD') {
    response.writeHead(405, { allow: 'GET, HEAD' }).end()
    return
  }
  if (url.pathname === '/office') {
    response.writeHead(302, { location: '/office/' }).end()
    return
  }
  const relative = url.pathname === '/office/' ? 'index.html' : decodeURIComponent(url.pathname.slice('/office/'.length))
  const path = resolve(root, ...relative.split('/'))
  if (!contained(relative, path) || !existsSync(path) || !statSync(path).isFile()) {
    response.writeHead(contained(relative, path) ? 404 : 403).end()
    return
  }
  response.writeHead(200, {
    'content-type': types.get(extname(path)) ?? 'application/octet-stream',
  })
  if (request.method === 'HEAD') {
    response.end()
    return
  }
  await pipeline(createReadStream(path), response)
}

function contained(relative, path) {
  if (relative.length === 0 || relative.split('/').some(part => part === '' || part === '.' || part === '..')) return false
  const base = resolve(root)
  return path === base || path.startsWith(`${base}${sep}`)
}
