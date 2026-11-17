import { encodeUniverfile } from '@univer/collab-gateway-contract'
import { VIEWER_BASE, VIEWER_WS_TUNNEL } from '../../shared/wire/viewer'
import { viewerStreamBaseUrl } from './desktop-transport.ts'

/** Server origin + the .univer being viewed (passed to WorktreeControlClient / viewer). */
export interface AppConfig {
  /** Same-origin (the vite dev server proxies `/uf` to the real ucb server). */
  origin: string
  /** Absolute .univer path being viewed, or the same-origin gateway route for display. */
  univerfile: string
  /** Same-origin gateway file key from `?file=<enc>`. */
  gatewayFileKey?: string
}

export type AppMode = 'standalone' | 'embedded'
export type AppContentScope = 'trunk' | 'worktree' | 'mergePreview'

/**
 * Product-comparison endpoint requested by an embedded host.
 *
 * The right panel embeds the Viewer with `mode=embedded`, which renders no shell chrome, so the
 * panel header owns the View/Diff switch and passes its intent through the address instead.
 */
export type AppComparisonSource = 'trunk' | { readonly worktreeId: string }

/** What the address bar encodes: which file, which worktree (absent = current version), which unit. */
export interface AppLocation {
  univerfile: string | null
  gatewayFileKey: string | null
  worktreeId: string | null
  unitId: string | null
  mode: AppMode
  scope: AppContentScope
  editable: boolean | null
  /** Diff mode requested by an embedded host; `null` keeps the single-version view. */
  comparison: AppComparisonSource | null
}

export interface WriteLocationOptions {
  univerfile: string
  gatewayFileKey?: string
  worktreeId?: string
  unitId?: string
  mode?: AppMode
  scope?: AppContentScope
  editable?: boolean
  /** Embedded comparison mode requested by the host panel. */
  view?: 'view' | 'diff'
  compare?: string
  /** Shell UI language; written only when the address bar carried it or the user toggled it. */
  lang?: string
}

/** Read the current view from the address bar. No file is a real "nothing selected" state. */
export function readLocation(): AppLocation {
  const p = new URLSearchParams(location.search)
  const mode = readAppMode(p)
  const file = p.get('file')
  const gatewayFileKey = file !== null && isGatewayFileKey(file) ? file : null
  return {
    univerfile: gatewayFileKey === null ? file : null,
    gatewayFileKey,
    worktreeId: p.get('worktree'),
    unitId: p.get('unit'),
    mode,
    scope: readContentScope(p, mode),
    editable: mode === 'embedded' ? readEditable(p) : null,
    comparison: readComparison(p, mode)
  }
}

/**
 * Reflect the current view in the address bar (`?file=…&worktree=…&unit=…`) so the user always
 * sees what they're looking at, refresh keeps it, and the URL is shareable. replaceState — we
 * don't push a history entry per click.
 */
export function writeLocation(loc: WriteLocationOptions): void {
  const p = new URLSearchParams()
  // Carried through address-bar rewrites so a refresh re-authorizes the same session scope.
  const sessionId = new URLSearchParams(location.search).get('sessionId')
  if (loc.gatewayFileKey !== undefined) {
    p.set('file', loc.gatewayFileKey)
  } else {
    p.set('file', loc.univerfile)
  }
  if (loc.mode === 'embedded') {
    p.set('mode', loc.mode)
    if (loc.scope !== undefined) {
      p.set('scope', loc.scope)
    }
    if (loc.editable !== undefined) {
      p.set('editable', loc.editable ? 'true' : 'false')
    }
    if (loc.view !== undefined) {
      p.set('view', loc.view)
      if (loc.view === 'diff' && loc.compare !== undefined) p.set('compare', loc.compare)
    }
  }
  if (loc.worktreeId !== undefined) {
    p.set('worktree', loc.worktreeId)
  }
  if (loc.unitId !== undefined) {
    p.set('unit', loc.unitId)
  }
  if (loc.lang !== undefined) {
    p.set('lang', loc.lang)
  }
  if (sessionId !== null) {
    p.set('sessionId', sessionId)
  }
  history.replaceState(null, '', `${location.pathname}?${p.toString()}`)
}

/**
 * True when this document is served by the DSH same-origin proxy (`/univer-viewer/...`) rather
 * than directly by the Gateway; only then do WebSocket endpoints need the exact-path tunnel.
 */
export function isProxyServedViewer(): boolean {
  const pathname = location.pathname
  return pathname === VIEWER_BASE || pathname.startsWith(`${VIEWER_BASE}/`)
}

/**
 * Rewrite a WebSocket endpoint through the fixed same-origin tunnel (`?target=<path>`) when
 * proxy-served: DSH upgrades register exact paths only, while the Gateway's `/uf/...` endpoints
 * are file-scoped and dynamic. Desktop uses the Host's explicit loopback transport address.
 */
export function resolveWebSocketUrl(url: string): string {
  if (!isProxyServedViewer()) return url
  const target = new URL(url, location.origin)
  // Endpoint input chooses only the Gateway path, never the network authority.
  const tunnel = new URL(VIEWER_WS_TUNNEL, viewerStreamBaseUrl())
  tunnel.searchParams.set('target', `${target.pathname}${target.search}`)
  tunnel.protocol = tunnel.protocol === 'https:' ? 'wss:' : 'ws:'
  return tunnel.toString()
}

/** Gateway file endpoint prefix for the configured target, including lifecycle WebSocket URLs. */
export function ufPrefix(config: AppConfig): string {
  if (config.gatewayFileKey !== undefined) {
    return `/uf/${config.gatewayFileKey}`
  }
  return `/uf/${encodeUniverfile(config.univerfile)}`
}

export function gatewayFileEndpointFromKey(origin: string, key: string): string {
  return `${origin.replace(/\/+$/u, '')}/uf/${key}`
}

function readAppMode(params: URLSearchParams): AppMode {
  return params.get('mode') === 'embedded' ? 'embedded' : 'standalone'
}

function readContentScope(params: URLSearchParams, mode: AppMode): AppContentScope {
  if (mode !== 'embedded') {
    return params.get('worktree') === null ? 'trunk' : 'worktree'
  }
  const scope = params.get('scope')
  if (scope === 'trunk' || scope === 'worktree' || scope === 'mergePreview') {
    return scope
  }
  return params.get('worktree') === null ? 'trunk' : 'worktree'
}

function readEditable(params: URLSearchParams): boolean | null {
  const editable = params.get('editable')
  if (editable === 'true') {
    return true
  }
  if (editable === 'false') {
    return false
  }
  return null
}

/**
 * Diff mode is an embedded-host request only. A standalone address keeps its own View/Diff control
 * in the shell, so carrying the parameter there would fight with that control on every rewrite.
 */
function readComparison(params: URLSearchParams, mode: AppMode): AppComparisonSource | null {
  if (mode !== 'embedded' || params.get('view') !== 'diff') return null
  const compare = params.get('compare')
  return compare === null || compare.length === 0 || compare === 'trunk'
    ? 'trunk'
    : { worktreeId: compare }
}

function isGatewayFileKey(value: string): boolean {
  return /^[A-Za-z0-9_-]+$/u.test(value)
}
