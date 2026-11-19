import { randomBytes } from 'node:crypto'
import { createReadStream, createWriteStream, existsSync, statSync, unlinkSync } from 'node:fs'
import { createServer } from 'node:http'
import { extname, join, normalize } from 'node:path'
import { pipeline } from 'node:stream/promises'
import { fileURLToPath } from 'node:url'
import { rename } from 'node:fs/promises'

const root = fileURLToPath(new URL('.', import.meta.url))
const port = Number(process.env.PORT ?? 4173)
const types = new Map([
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.css', 'text/css; charset=utf-8'],
  ['.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation']
])

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? '/', `http://127.0.0.1:${port}`)
    if (request.method === 'POST' && url.pathname === '/save') {
      await save(url.searchParams.get('name') ?? '', request, response)
      return
    }
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      response.writeHead(405).end()
      return
    }
    const name = decodeURIComponent(url.pathname)
    const relative = name === '/' ? 'index.html' : name.replace(/^\/+/, '')
    if (!safe(relative)) {
      response.writeHead(403).end()
      return
    }
    const path = join(root, relative)
    if (!existsSync(path) || !statSync(path).isFile()) {
      response.writeHead(404).end()
      return
    }
    response.writeHead(200, {
      'content-type': types.get(extname(path)) ?? 'application/octet-stream'
    })
    if (request.method === 'HEAD') {
      response.end()
      return
    }
    await pipeline(createReadStream(path), response)
  } catch {
    if (!response.headersSent) response.writeHead(500)
    response.end()
  }
})

server.listen(port, '127.0.0.1')

function safe(name) {
  if (name.length === 0 || name.includes('..')) return false
  const normalized = normalize(name)
  return normalized === name && !normalized.startsWith('..')
}

async function save(name, request, response) {
  if (!safe(name) || extname(name) !== '.pptx') {
    response.writeHead(400).end()
    return
  }
  const path = join(root, name)
  if (!existsSync(path) || !statSync(path).isFile()) {
    response.writeHead(404).end()
    return
  }
  const temporary = `${path}.${randomBytes(4).toString('hex')}.tmp`
  try {
    await pipeline(request, createWriteStream(temporary))
    await rename(temporary, path)
  } catch (error) {
    try {
      unlinkSync(temporary)
    } catch {
      // The temp file was never created.
    }
    throw error
  }
  response.writeHead(204).end()
}
