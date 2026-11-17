// Built Client half in jsdom: real React, a mock DSH ctx, and a fake /univer-api HTTP server.
// Covers the single right-Sidebar preview (embedded Viewer, version switching, View/Diff
// comparison, ready / reopen / merge / discard) as the only Univer surface, the write-driven reveal
// that keeps agent work inside that one tab, the docked fallback for hosts without a Sidebar, and
// the real Cordis lifecycle.
//
// The conversation log itself is asserted to stay free of Univer surfaces: the preview tab is the
// plugin's only rendering surface, and the input seat either reveals it or (on a host without a
// right Sidebar) renders the docked fallback. Scenarios run in order and share the mounted DOM,
// like a user session.
import { readFileSync } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'
import { closeServer, listen } from '../../support/net.ts'
import { pluginRoot, repoPath } from '../../support/paths.ts'

const repoRequire = createRequire(repoPath('package.json'))

const SESSION_ID = 'test-session-id'
const DEMO_FILE = join(tmpdir(), 'dsh-univer-client-smoke', 'demo.univer')
const SECOND_FILE = join(tmpdir(), 'dsh-univer-client-smoke', 'second.univer')
const SESSION_CWD = join(tmpdir(), 'dsh-univer-client-smoke', 'workdir')
const REL_DEMO_FILE = join(SESSION_CWD, 'work_班级成绩表', '班级管理.univer')
const WINDOWS_CWD = 'C:\\Users\\17361\\Documents\\DSH'
const WINDOWS_FILE = WINDOWS_CWD + '\\学生成绩表.univer'
const WORKTREE = 'wt-msvqmweb-47hcdg'
const OTHER_WORKTREE = 'wt-other0000-aaaaaa'
const VIEW_URL =
  'http://127.0.0.1:9123/?file=KEY&worktree=wt-msvqmweb-47hcdg&mode=embedded&scope=worktree'
const TRUNK_URL = 'http://127.0.0.1:9123/?file=KEY'
const withLang = (url: string, lang: string) => {
  const target = new URL(url)
  target.searchParams.set('lang', lang)
  return target.toString()
}
const UNITS = [
  {
    unitId: 'u-msvo3wpe-p4pqi4',
    name: '销售',
    type: 'sheet',
    kind: 'modified',
    worktreeUrl: VIEW_URL + '&unit=u-msvo3wpe-p4pqi4'
  },
  {
    unitId: 'u-msvy1lry-dv3hia',
    name: '班级成绩汇报',
    type: 'slide',
    kind: 'added',
    worktreeUrl: VIEW_URL + '&unit=u-msvy1lry-dv3hia'
  }
]
const DEFAULT_UNIT_URL = UNITS[0]!.worktreeUrl
const SLIDE_UNIT_URL = UNITS[1]!.worktreeUrl
const ZH_DEFAULT_UNIT_URL = withLang(DEFAULT_UNIT_URL, 'zh-CN')
const EN_SLIDE_UNIT_URL = withLang(SLIDE_UNIT_URL, 'en-US')
/** The trunk target the panel scopes itself. */
const trunkUrl = () => withLang(TRUNK_URL + '&mode=embedded&scope=trunk', 'zh-CN')
const diffUrl = (base: string) => {
  const target = new URL(base)
  target.searchParams.set('view', 'diff')
  return withLang(target.toString(), 'zh-CN')
}
const SESSION_ADDRESS = 'dsh-resource://file/session/test-session-id/预览演示/月度支出表.univer'

/**
 * The one branch the Host projects for a file.
 *
 * `unsaved` is what enables Save and Discard: it is true while the branch still differs from the
 * last saved version, and the fake Host clears the branch once either decision settles it.
 */
const branchState = (unsaved = true, worktreeId = WORKTREE) => ({
  worktreeId,
  name: worktreeId === WORKTREE ? 'v3smoke' : 'other',
  unsaved,
  worktreeUrl: VIEW_URL,
  units: UNITS
})

/** Encode one path the way the file-resource grammar encodes it. */
const encodeAddressPath = (path: string) =>
  path
    .split(/[\\/]/)
    .map((segment) => encodeURIComponent(segment).replace(/%3A/gi, ':'))
    .join('/')

/** Compare two file paths regardless of the platform's separator. */
const samePath = (left: string, right: string) =>
  left.replaceAll('\\', '/') === right.replaceAll('\\', '/')

export function defineClientSuite(): void {
  describe('built Client', () => {
    // ---- fake Host state ----
    let branch: any = null
    const missingFiles = new Set<string>()
    const stateRequests: Array<string | null> = []
    /** Bindings the docked panel asked the Host to hold for this session. */
    const panelOpens: Array<Record<string, unknown>> = []
    let panelCloses = 0
    /** Lifecycle decisions the dock header submitted. */
    const actionRequests: Array<Record<string, unknown>> = []
    /** When set, the next lifecycle request is rejected with this reason. */
    let actionFailure: string | undefined
    /** Gateway phase the fake Host reports; a stopped Gateway cannot project Viewer URLs. */
    let gatewayPhase = 'running'
    const currentState = () => ({
      ok: true,
      file: DEMO_FILE,
      gateway: 'http://127.0.0.1:9123',
      gatewayRunning: true,
      viewerUrl: TRUNK_URL,
      branch
    })

    // ---- mock ctx state ----
    const slotEntries: Array<{ options: any; Component: any }> = []
    let localeDicts: { ns: string; dicts: Record<string, Record<string, string>> } | null = null
    let conversationDefinition: any = null
    let activeLocale = 'zh'
    let localeRevision = 0
    const localeListeners = new Set<() => void>()
    /** File viewers the Client registered with the optional sidebar service. */
    const sidebarViewers: any[] = []
    /** Native right-Sidebar tab types the Client registered. */
    const nativeTabTypes: any[] = []
    /**
     * The right-Sidebar navigator, or `undefined` for a host that has no right Sidebar at all.
     * The dock tests run without it (that host renders the fallback panel); the reveal tests
     * install one and record the addresses it was asked to open.
     */
    let sidebarNavigator:
      | { readonly opened: string[]; openResource(address: string): void }
      | undefined = undefined
    /** Install a navigator that records every address the Client asks to reveal. */
    const installSidebarNavigator = (): string[] => {
      const opened: string[] = []
      sidebarNavigator = {
        opened,
        openResource(address: string) {
          opened.push(address)
        }
      }
      return opened
    }
    /** The target the fake Host reports as this session's binding, for `/univer-api/panel`. */
    let panelBinding: Record<string, unknown> | null = null
    const effects: unknown[] = []
    const conversationEventRegistry = {
      register(definition: any) {
        conversationDefinition = definition
        return () => {
          conversationDefinition = null
        }
      }
    }
    /** Switch the active locale the way the runtime does: bump the revision and notify. */
    const setActiveLocale = (next: string) => {
      activeLocale = next
      localeRevision += 1
      for (const listener of localeListeners) listener()
    }
    const fakeCtx: any = {
      inject(services: string[], callback: (ctx: any) => unknown) {
        const key = services.join(',')
        if (key !== 'betterSidebar' && key !== 'sidebarRightTabs')
          throw new Error(`unexpected ctx.inject(${JSON.stringify(services)})`)
        return callback(fakeCtx)
      },
      effect(fn: () => unknown) {
        const disposer = fn()
        effects.push(disposer)
        return () => {
          if (typeof disposer === 'function') disposer()
        }
      },
      slots: {
        register(options: any, Component: any) {
          slotEntries.push({ options, Component })
          return () => {}
        },
        inject(key: string, callback: () => unknown) {
          if (key !== 'conversation.input.dock' && key !== 'sidebar.right.pane.tab')
            throw new Error(`unexpected slots.inject("${key}")`)
          return callback()
        }
      },
      locale: {
        register(ns: string, dicts: Record<string, Record<string, string>>) {
          localeDicts = { ns, dicts }
          return () => {}
        },
        bind() {
          return (key: string) => localeDicts?.dicts[activeLocale]?.[key] ?? key
        },
        getSnapshot() {
          return { active: activeLocale, revision: localeRevision }
        },
        subscribe(listener: () => void) {
          localeListeners.add(listener)
          return () => localeListeners.delete(listener)
        }
      },
      get(name: string) {
        if (name === 'uiConversation') return { events: conversationEventRegistry }
        if (name === 'sidebarRight') return sidebarNavigator
        if (name === 'sidebarRightTabs')
          return {
            register(definition: any) {
              nativeTabTypes.push(definition)
              return () => {}
            }
          }
        if (name === 'betterSidebar')
          return {
            registerFileViewer(descriptor: any) {
              sidebarViewers.push(descriptor)
              return () => {}
            }
          }
        throw new Error(`unexpected ctx.get("${name}")`)
      }
    }

    // ---- harness assigned in beforeAll ----
    let server: Server
    let dom: any
    let document: Document
    let React: any
    let createRoot: (element: Element) => any
    let pluginExports: any
    let dockEntry: any
    let dockInjected: any
    let fileViewer: any
    let nativeTab: any
    let nativeBodyEntry: any
    let callSequence = 0

    const t = (key: string) => localeDicts?.dicts[activeLocale]?.[key] ?? key
    const q = (selector: string): any => document.querySelector(selector)
    const qa = (selector: string): any[] => Array.from(document.querySelectorAll(selector))
    const click = (element: Element) =>
      element.dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }))
    const pointer = (element: Element, type: string, init: Record<string, unknown>) =>
      element.dispatchEvent(new dom.window.PointerEvent(type, { bubbles: true, ...init }))
    const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))
    async function waitFor(description: string, predicate: () => boolean, timeoutMs = 5000) {
      const startedAt = Date.now()
      while (Date.now() - startedAt < timeoutMs) {
        if (predicate()) return
        await sleep(80)
      }
      throw new Error(
        `timeout waiting for: ${description}\nhtml: ${document.body.innerHTML.slice(0, 1500)}`
      )
    }

    const operation = (
      name: string,
      worktreeId: string | null,
      action: string | null = null,
      unitId: string | null = null,
      phase = 'succeeded'
    ) => ({
      callId: `fixture-${++callSequence}`,
      name,
      action,
      file: DEMO_FILE,
      worktreeId,
      unitId,
      phase
    })
    const turnFile = (
      file: string,
      worktreeId: string | null = null,
      name = worktreeId === null ? 'status' : 'execute',
      action: string | null = null,
      phase = 'succeeded',
      unitId: string | null = null
    ) => ({ file, operations: [{ ...operation(name, worktreeId, action, unitId, phase), file }] })
    const sessionWithFiles = (files: any[], running: boolean) => ({
      sessionId: SESSION_ID,
      running,
      chat: {
        timeline: {
          turns: new Map([
            [3, { data: { get: (key: string) => (key === 'univerTurn' ? { files } : undefined) } }]
          ])
        }
      }
    })
    const sessionWithTargets = (
      targets: Array<{ file: string; worktreeId: string | null; name?: string }>,
      running = true
    ) =>
      sessionWithFiles(
        targets.map((target) => turnFile(target.file, target.worktreeId, target.name)),
        running
      )
    const runtimeProps = (session: any) => ({
      session: { sessionId: session.sessionId, running: session.running },
      useSession: (selector: any) =>
        selector({ sessionId: session.sessionId, running: session.running }),
      useChat: (selector: any) => selector(session.chat)
    })
    const useSessionsAt = (cwd: string) => (selector: any) =>
      selector({ byId: { [SESSION_ID]: { cwd } } })
    const mount = (tag = 'div') => {
      const element = document.createElement(tag)
      document.body.appendChild(element)
      return { element, root: createRoot(element) }
    }
    /** Props the input-dock seat hands the Univer dock. */
    const dockProps = (session: any, cwd = SESSION_CWD) => ({
      ...runtimeProps(session),
      t,
      getViewerLocale: dockInjected.getViewerLocale,
      hasSidebarPreview: dockInjected.hasSidebarPreview,
      openPreview: dockInjected.openPreview,
      sessionId: SESSION_ID,
      useSessions: useSessionsAt(cwd)
    })
    const panelFrame = () => q('.uvf_dockFrame')?.getAttribute('src')
    const panelChip = () => q('.uvf_dockChip')?.textContent

    beforeAll(async () => {
      server = createServer(async (req, res) => {
        const url = new URL(req.url ?? '/', 'http://x')
        const json = (status: number, body: unknown) => {
          res.writeHead(status, { 'content-type': 'application/json' })
          res.end(JSON.stringify(body))
        }
        const readBody = async () => {
          const chunks: Buffer[] = []
          for await (const chunk of req) chunks.push(chunk)
          return JSON.parse(Buffer.concat(chunks).toString('utf8'))
        }
        if (req.method === 'GET' && url.pathname === '/univer-api/status') {
          json(200, {
            gateway: {
              phase: gatewayPhase,
              gateway: gatewayPhase === 'running' ? 'http://127.0.0.1:9123' : null,
              owned: false
            },
            unitContent: 'bundled'
          })
          return
        }
        if (req.method === 'GET' && url.pathname === '/univer-api/state') {
          const file = url.searchParams.get('file')
          if (url.searchParams.get('sessionId') !== SESSION_ID) {
            res.writeHead(400).end()
            return
          }
          stateRequests.push(file)
          if (gatewayPhase !== 'running') {
            json(503, { ok: false, code: 'GATEWAY_UNAVAILABLE', message: 'no gateway' })
            return
          }
          if (file !== null && missingFiles.has(file)) {
            json(400, { ok: false, code: 'INVALID_FILE_PATH', message: 'path does not exist' })
            return
          }
          if (
            file === null ||
            ![DEMO_FILE, REL_DEMO_FILE, SECOND_FILE].some((known) => samePath(file, known))
          ) {
            res.writeHead(404).end()
            return
          }
          json(200, currentState())
          return
        }
        if (req.method === 'POST' && url.pathname === '/univer-api/worktree-action') {
          const body = await readBody()
          if (body.sessionId !== SESSION_ID) {
            res.writeHead(400).end()
            return
          }
          actionRequests.push(body)
          if (actionFailure !== undefined) {
            json(200, {
              ok: false,
              action: body.action,
              worktreeId: body.worktreeId,
              reason: actionFailure
            })
            return
          }
          // Both decisions settle the branch: Save folds it into the saved version, Discard drops
          // it. Either way the file then has no unsaved changes to show.
          if (branch !== null && branch.worktreeId === body.worktreeId) branch = null
          json(200, {
            ok: true,
            action: body.action,
            worktreeId: body.worktreeId,
            state: currentState()
          })
          return
        }
        if (req.method === 'POST' && url.pathname === '/univer-api/panel/open') {
          const body = await readBody()
          if (body.sessionId !== SESSION_ID) {
            res.writeHead(400).end()
            return
          }
          panelOpens.push(body)
          // A stopped Gateway cannot authorize or import anything, exactly like the Host.
          if (gatewayPhase !== 'running') {
            json(503, { ok: false, code: 'GATEWAY_UNAVAILABLE', message: 'no gateway' })
            return
          }
          const requested = typeof body.file === 'string' ? body.file : ''
          if (requested !== '' && missingFiles.has(requested)) {
            json(400, { ok: false, code: 'INVALID_FILE_PATH', message: 'path does not exist' })
            return
          }
          // A `.pptx` is imported into the sibling `.univer`; either way the binding names the
          // container, which is what the preview then polls.
          const container = requested.replace(/\.pptx$/i, '.univer')
          const presentation = container !== requested
          // The Host only binds a live branch: a file whose changes were saved or discarded has
          // nothing open to bind and answers WORKTREE_NOT_FOUND.
          const requestedWorktree =
            typeof body.worktreeId === 'string' && body.worktreeId !== ''
              ? body.worktreeId
              : undefined
          const worktreeId = requestedWorktree ?? branch?.worktreeId
          if (worktreeId === undefined && !presentation) {
            json(404, {
              ok: false,
              code: 'WORKTREE_NOT_FOUND',
              message: 'No draft worktree holds content for this file.'
            })
            return
          }
          const target = branch?.worktreeId === worktreeId ? branch : undefined
          const unitId =
            typeof body.unitId === 'string' && body.unitId !== ''
              ? body.unitId
              : (target?.units[0]?.unitId ?? '')
          json(200, {
            ok: true,
            panel: {
              source: presentation ? requested : null,
              file: container,
              worktreeId: worktreeId ?? '',
              unitId,
              viewerUrl: ''
            }
          })
          return
        }
        if (req.method === 'POST' && url.pathname === '/univer-api/panel/close') {
          const body = await readBody()
          if (body.sessionId !== SESSION_ID) {
            res.writeHead(400).end()
            return
          }
          panelCloses += 1
          json(200, { ok: true })
          return
        }
        if (req.method === 'GET' && url.pathname === '/univer-api/panel') {
          json(200, { ok: true, panel: panelBinding })
          return
        }
        res.writeHead(404).end()
      })
      const origin = `http://127.0.0.1:${await listen(server)}`

      const { JSDOM } = repoRequire('jsdom')
      dom = new JSDOM('<!doctype html><html><head></head><body></body></html>', {
        url: origin + '/'
      })
      // jsdom does not implement PointerEvent; the splitter drag needs it.
      if (dom.window.PointerEvent === undefined) {
        dom.window.PointerEvent = class PointerEvent extends dom.window.MouseEvent {
          pointerId: number
          pointerType: string
          isPrimary: boolean
          constructor(type: string, params: any = {}) {
            super(type, params)
            this.pointerId = params.pointerId ?? 0
            this.pointerType = params.pointerType ?? 'mouse'
            this.isPrimary = params.isPrimary ?? true
          }
        }
      }
      // jsdom implements neither pointer capture nor the dialog top layer; the panel only needs
      // the calls to exist, and containment is verified in the browser layout test.
      dom.window.Element.prototype.setPointerCapture ??= function () {}
      dom.window.Element.prototype.releasePointerCapture ??= function () {}
      dom.window.Element.prototype.hasPointerCapture ??= function () {
        return false
      }
      document = dom.window.document
      Object.assign(globalThis, { window: dom.window, document })
      Object.defineProperty(globalThis, 'navigator', {
        value: dom.window.navigator,
        configurable: true
      })
      Object.defineProperty(dom.window, 'innerWidth', {
        value: 1440,
        writable: true,
        configurable: true
      })
      Object.defineProperty(dom.window, 'innerHeight', {
        value: 1000,
        writable: true,
        configurable: true
      })

      React = repoRequire('react')
      const jsxRuntime = repoRequire('react/jsx-runtime')
      ;({ createRoot } = repoRequire('react-dom/client'))
      dom.window.__ModuleLoader__ = {
        load({ factory }: { factory: (require: (spec: string) => unknown) => unknown }) {
          pluginExports = factory((spec) => {
            if (spec === 'react') return React
            if (spec === 'react/jsx-runtime') return jsxRuntime
            throw new Error(`unexpected require("${spec}")`)
          })
        }
      }
      const source = readFileSync(join(pluginRoot(), 'lib/client.js'), 'utf8')
      new Function('window', `${source}\n//# sourceURL=lib/client.js`)(dom.window)
    })

    afterAll(async () => {
      if (server !== undefined) await closeServer(server)
      dom?.window.close()
      for (const key of ['window', 'document', 'navigator']) Reflect.deleteProperty(globalThis, key)
    })

    describe('registration', () => {
      it('registers through the window ModuleLoader', () => {
        expect(
          pluginExports,
          'client module did not register via __ModuleLoader__.load'
        ).not.toBeNull()
        expect(pluginExports.apply).toEqual(expect.any(Function))
        expect(pluginExports.inject, 'Conversation services of the supported DSH line').toEqual([
          'slots',
          'locale',
          'conversation',
          'uiConversation'
        ])
        pluginExports.apply(fakeCtx)
      })

      it('registers one conversation seat that reveals the single preview', () => {
        dockEntry = slotEntries.find(
          (entry) =>
            entry.options.name === 'conversation.input.dock' && entry.options.id === 'univer-dock'
        )
        expect(dockEntry, 'dock entry').toBeDefined()
        // The conversation log stays pure conversation: no Turn-tail card, and no settings card
        // because the plugin no longer has settings switches.
        expect(slotEntries.map((entry) => entry.options.name).sort()).toEqual([
          'conversation.input.dock',
          'sidebar.right.pane.tab'
        ])
        expect(dockEntry.options.locale).toBe('univer')
        expect(localeDicts?.ns).toBe('univer')
        expect(conversationDefinition?.kind).toBe('univerTurn')
        dockInjected = dockEntry.options.inject()
        expect(dockInjected.getViewerLocale).toEqual(expect.any(Function))
        // The seat holds no UI of its own on a host with a right Sidebar: it reveals the preview.
        expect(dockInjected.hasSidebarPreview).toEqual(expect.any(Function))
        expect(dockInjected.openPreview).toEqual(expect.any(Function))
      })

      // A click on the file in the DSH file tree must reach the Viewer with no agent write, so
      // the Client registers a previewer for the `.univer` extension.
      it('registers an on-demand .univer file viewer', () => {
        expect(sidebarViewers).toHaveLength(1)
        fileViewer = sidebarViewers[0]
        expect(fileViewer).toMatchObject({ id: 'univer-office:univer', exts: ['univer', 'pptx'] })
        // It must outrank the sidebar catch-all viewer.
        expect(fileViewer.priority).toBeGreaterThan(0)
        // `.univer` is a binary container: a byte-reading strategy would render the download pane.
        expect(fileViewer.fetchStrategy).toBe('none')
        expect(fileViewer.component).toEqual(expect.any(Function))
      })

      // dsh-better-sidebar wins the address claim whenever it is installed, so the registration
      // above only covers hosts that have it. Without this one, a host on DSH's own file tree has
      // no `.univer` preview at all.
      it('registers a native right-Sidebar tab type', () => {
        expect(nativeTabTypes).toHaveLength(1)
        nativeTab = nativeTabTypes[0]
        expect(nativeTab).toMatchObject({
          id: 'dsh-univer-office',
          kind: 'univer',
          patterns: ['*.univer', '*.pptx'],
          // A third-party tab type claims addresses in the highest band.
          priority: 'extension'
        })
        expect(nativeTab.canOpen(SESSION_ADDRESS)).toBe(true)
        // The absolute scope carries no session, so the state route cannot authorize it.
        expect(nativeTab.canOpen('dsh-resource://file/absolute/home/me/x.univer')).toBe(false)
        expect(nativeTab.canOpen('dsh-resource://note/session/s1/x.univer')).toBe(false)
        expect(nativeTab.title(SESSION_ADDRESS)).toBe('月度支出表.univer')
        nativeBodyEntry = slotEntries.find(
          (entry) =>
            entry.options.name === 'sidebar.right.pane.tab' && entry.options.key === nativeTab.id
        )
        expect(nativeBodyEntry?.Component).toEqual(expect.any(Function))
      })

      // The mirrored `dsh-resource://file/…` grammar, exercised through the two members the
      // sidebar actually calls rather than by reaching into the parser.
      it.each([
        ['dsh-resource://file/session/s1/src/a.univer', true, 'a.univer'],
        ['dsh-resource://file/session/s1/%E6%9C%88%E5%BA%A6.univer', true, '月度.univer'],
        ['dsh-resource://file/session/s1/a%20b.univer', true, 'a b.univer'],
        ['dsh-resource://file/session/s1/a.univer?line=3#frag', true, 'a.univer'],
        ['dsh-resource://file/session/s1/C:/dir/a.univer', true, 'a.univer'],
        ['dsh-resource://file/absolute/home/me/a.univer', false, undefined],
        ['dsh-resource://note/session/s1/a.univer', false, undefined],
        ['dsh-resource://file/session/s1/%ZZ.univer', false, undefined],
        ['sidebar://guide', false, undefined],
        ['/plain/path/a.univer', false, undefined]
      ])('claims %s: %s', (address, claimable, title) => {
        expect(nativeTab.canOpen(address)).toBe(claimable)
        if (title !== undefined) expect(nativeTab.title(address)).toBe(title)
      })
    })

    describe('Conversation definition', () => {
      const location = { kind: 'turn', turn: 7 }
      const mkContext = (state: any) => ({
        state,
        key: '',
        kind: 'univerTurn',
        id: '7',
        matches: [],
        start: undefined,
        current: new Map()
      })
      const startMatch = {
        id: '7',
        role: 'start',
        event: { type: 'turn/start', data: { turn: 7 } },
        location
      }
      const toolCall = (callId: string, name: string, args: unknown) => ({
        id: '7',
        role: 'update',
        location,
        event: {
          type: 'tool/call',
          data: { turn: 7, step: 1, callId, name, arguments: JSON.stringify(args) }
        }
      })
      const toolResult = (message: unknown) => ({
        id: '7',
        role: 'update',
        location,
        event: { type: 'tool/result', data: { turn: 7, step: 1, message } }
      })
      const readyCall = toolCall('call-ready', 'univer_worktree', {
        action: 'discard',
        file: '/x/proj/notes/demo.univer',
        worktreeId: 'wt-abc12345'
      })
      const readyText = {
        type: 'text',
        text: JSON.stringify({
          operation: 'worktree',
          file: '/x/proj/notes/demo.univer',
          result: { action: 'discard', worktreeId: 'wt-abc12345' }
        })
      }
      const start = () =>
        conversationDefinition.start({ state: undefined }, startMatch, {
          previous: () => undefined
        })

      it('starts an empty Turn', () => {
        expect(start()).toMatchObject({ turn: 7, files: [] })
      })

      // Session format v4 (DSH 0.1.7): the tool result is a first-class tool-role message —
      // `toolCallId` / `isError` sit on the message and `content` holds the text blocks directly.
      // 0.3.2 threw "reading 'flatMap'" on this shape and took the whole timeline down with it.
      it('applies a v4 tool-role result and ignores a foreign one', () => {
        let state = conversationDefinition.update(mkContext(start()), readyCall)
        state = conversationDefinition.update(
          mkContext(state),
          toolResult({
            role: 'tool',
            source: { kind: 'tool', callId: 'call-ready' },
            toolCallId: 'call-ready',
            isError: false,
            content: [readyText]
          })
        )
        expect(state.files[0]?.operations).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              callId: 'call-ready',
              phase: 'succeeded',
              worktreeId: 'wt-abc12345'
            })
          ])
        )
        const foreign = conversationDefinition.update(
          mkContext(state),
          toolResult({
            role: 'tool',
            source: { kind: 'tool', callId: 'call-other' },
            toolCallId: 'call-other',
            content: [{ type: 'text', text: 'total 12\n-rw-r--r-- x' }]
          })
        )
        expect(foreign).toBe(state)
      })

      it('keeps branch actions across later reads and replays screenshots structurally', () => {
        const update = (state: any, match: any) =>
          conversationDefinition.update(mkContext(state), match)
        let state = update(start(), readyCall)
        state = update(
          state,
          toolResult({
            content: [{ type: 'tool-result', toolCallId: 'call-ready', content: [readyText] }]
          })
        )
        state = update(
          state,
          toolCall('call-status', 'univer_status', { file: '/x/proj/notes/demo.univer' })
        )
        state = update(
          state,
          toolCall('call-screenshot', 'univer_screenshot', {
            file: '/x/proj/notes/demo.univer',
            worktreeId: 'wt-abc12345',
            unitId: 'unit-1',
            output: 'shots'
          })
        )
        state = update(
          state,
          toolResult({
            content: [
              {
                type: 'tool-result',
                toolCallId: 'call-screenshot',
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      operation: 'screenshot',
                      file: '/x/proj/notes/demo.univer',
                      result: { unitId: 'unit-1', images: [{ name: 'page-1.png' }] }
                    })
                  },
                  { type: 'image', attachment: { attachmentId: 'fixture-image' } }
                ]
              }
            ]
          })
        )
        expect(state.files[0].operations).toHaveLength(3)
        expect(state.files[0].operations[0].action).toBe('discard')
        expect(state.files[0].operations[2].name).toBe('screenshot')
        const locationData = conversationDefinition.buildLocationData(mkContext(state), 'turn')
        expect(locationData?.key).toBe('univerTurn')
        expect(locationData?.value.files).toHaveLength(1)
      })
    })

    // The docked fallback is what a host without a right Sidebar shows. Each scenario starts from a
    // conversation snapshot and drives a freshly mounted panel, the way a user session does.
    describe('docked fallback panel', () => {
      let panel: { element: HTMLElement; root: any }
      const renderPanel = (session: any, cwd = SESSION_CWD) =>
        panel.root.render(React.createElement(dockEntry.Component, dockProps(session, cwd)))

      beforeEach(() => {
        branch = branchState(true)
        actionFailure = undefined
        gatewayPhase = 'running'
        panelOpens.length = 0
        actionRequests.length = 0
        panelCloses = 0
        panel = mount()
      })

      afterEach(async () => {
        panel.root.unmount()
        panel.element.remove()
        // Drain the release request the unmount may still be sending.
        await sleep(120)
      })

      // No Turn-tail registration exists, so the conversation log never renders a card; the docked
      // panel appears outside the chat even though the seat itself is inside the input dock.
      it('renders one panel outside the conversation and never a chat card', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('右侧面板出现', () => q('.uvf_dock') !== null)
        expect(q('.uvf_dock').parentElement).toBe(document.body)
        expect(qa('.uvf_panel')).toHaveLength(0)
        expect(qa('.uvf_win')).toHaveLength(0)
        expect(q('.uvf_dockFrame')?.getAttribute('src')).toBe(ZH_DEFAULT_UNIT_URL)
        expect(q('.uvf_dock')?.getAttribute('data-status')).toBe('unsaved')
        expect(panelChip()).toBe(t('dock.unsaved'))
        expect(q('.uvf_dock')?.getAttribute('data-file')).toBe(DEMO_FILE)
      })

      it('holds the displayed Unit as the session target', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('面板把当前 Unit 绑定到会话', () => panelOpens.length === 1)
        expect(panelOpens[0]).toMatchObject({
          file: DEMO_FILE,
          sessionId: SESSION_ID,
          worktreeId: WORKTREE,
          unitId: UNITS[0]!.unitId
        })
      })

      it('merges a relative tool-call path and its absolute result into one panel', async () => {
        renderPanel(
          sessionWithFiles(
            [turnFile('work_班级成绩表/班级管理.univer'), turnFile(REL_DEMO_FILE, WORKTREE)],
            false
          )
        )
        await waitFor(
          '相对路径和绝对路径去重为一个面板',
          () =>
            qa('.uvf_dock').length === 1 &&
            q('.uvf_dock')?.getAttribute('data-file') === REL_DEMO_FILE
        )
      })

      it('merges Win32 paths written with opposite separators into one panel', async () => {
        renderPanel(
          sessionWithFiles(
            [turnFile('学生成绩表.univer'), turnFile(WINDOWS_FILE.replaceAll('\\', '/'), WORKTREE)],
            false
          ),
          WINDOWS_CWD
        )
        await waitFor(
          'Windows 分隔符去重为一个面板',
          () => q('.uvf_dock')?.getAttribute('data-file') === WINDOWS_FILE
        )
      })

      it('shows the newest file when one Turn touches several, and no copy of the chat', async () => {
        renderPanel(
          sessionWithFiles([turnFile(DEMO_FILE, WORKTREE), turnFile(SECOND_FILE, WORKTREE)], false)
        )
        await waitFor(
          '最新目标接管面板',
          () => q('.uvf_dock')?.getAttribute('data-file') === SECOND_FILE
        )
        expect(qa('.uvf_dock')).toHaveLength(1)
      })

      // Reads never open the panel, and an already-open panel survives later reads.
      it('opens no panel for a read-only Turn', async () => {
        const isolated = mount()
        try {
          isolated.root.render(
            React.createElement(
              dockEntry.Component,
              dockProps(
                sessionWithTargets([{ file: DEMO_FILE, worktreeId: null, name: 'status' }], false)
              )
            )
          )
          await sleep(400)
          expect(q('.uvf_dock'), 'reads never open the dock').toBeNull()
          expect(panelOpens, 'a read never binds a Unit').toHaveLength(0)
        } finally {
          isolated.root.unmount()
          isolated.element.remove()
        }
      })

      it('offers the Unit chips and switches the Viewer without dropping the panel', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('单元选择出现', () => qa('.uvf_unit').length === 2)
        click(qa('.uvf_unit')[1])
        await waitFor(
          '切换单元后 Viewer 跟随',
          () => panelFrame() === withLang(SLIDE_UNIT_URL, 'zh-CN')
        )
        expect(qa('.uvf_dock')).toHaveLength(1)
      })

      it('follows a locale switch in place', async () => {
        const session = sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false)
        renderPanel(session)
        await waitFor('中文 Viewer', () => panelFrame() === ZH_DEFAULT_UNIT_URL)
        click(qa('.uvf_unit')[1])
        await waitFor('中文幻灯片单元', () => panelFrame() === withLang(SLIDE_UNIT_URL, 'zh-CN'))
        setActiveLocale('en')
        try {
          renderPanel(session)
          await waitFor('英文单元 Viewer', () => panelFrame() === EN_SLIDE_UNIT_URL)
          expect(qa('.uvf_dock')).toHaveLength(1)
        } finally {
          setActiveLocale('zh')
          renderPanel(session)
          await waitFor('切回中文 Viewer', () => panelFrame() === withLang(SLIDE_UNIT_URL, 'zh-CN'))
        }
      })

      it('offers Save and Discard only while the branch has unsaved changes', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('保存按钮出现', () => q('[data-panel-action=save]') !== null)
        expect(q('[data-panel-action=save]')?.disabled).toBe(false)
        expect(q('[data-panel-action=discard]')?.disabled).toBe(false)
        expect(panelChip()).toBe(t('dock.unsaved'))
        expect(q('.uvf_dock')?.getAttribute('data-status')).toBe('unsaved')
      })

      it('compares the branch against the saved version and back', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('对比按钮出现', () => q('[data-panel-action=compare]') !== null)
        click(q('[data-panel-action=compare]'))
        await waitFor('进入 Diff 视图', () => panelFrame() === diffUrl(DEFAULT_UNIT_URL))
        expect(q('[data-panel-action=compare]')?.getAttribute('aria-pressed')).toBe('true')
        // There is exactly one named baseline — the last saved version — so no picker is offered.
        expect(qa('select[aria-label="对比基准"]')).toHaveLength(0)
        click(q('[data-panel-action=compare]'))
        await waitFor('退出 Diff 视图', () => panelFrame() === ZH_DEFAULT_UNIT_URL)
      })

      it('saves the branch into the last saved version from the toolbar', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('保存按钮出现', () => q('[data-panel-action=save]') !== null)
        click(q('[data-panel-action=save]'))
        await waitFor('保存后回到已保存状态', () => panelChip() === t('dock.saved'))
        expect(actionRequests.at(-1)).toMatchObject({ action: 'merge', worktreeId: WORKTREE })
        expect(q('[data-panel-action=save]')?.disabled).toBe(true)
        expect(q('[data-panel-action=discard]')?.disabled).toBe(true)
        await waitFor('保存后展示当前版本', () => panelFrame() === trunkUrl())
      })

      it('discards the branch only after the confirmation', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('丢弃按钮出现', () => q('[data-panel-action=discard]') !== null)
        click(q('[data-panel-action=discard]'))
        await waitFor('出现丢弃确认', () => q('[data-panel-action=confirm-discard]') !== null)
        expect(q('[role=alertdialog]')?.textContent).toContain(t('dock.confirmDiscard'))
        expect(actionRequests, 'opening the dialog submits nothing').toHaveLength(0)
        click(q('[data-panel-action=confirm-discard]'))
        await waitFor('丢弃后回到已保存状态', () => panelChip() === t('dock.saved'))
        expect(actionRequests.at(-1)).toMatchObject({ action: 'discard', worktreeId: WORKTREE })
        await waitFor('丢弃后展示当前版本', () => panelFrame() === trunkUrl())
        expect(q('[data-panel-action=discard]')?.disabled).toBe(true)
      })

      it('reports a rejected action without losing the unsaved branch', async () => {
        actionFailure = 'worktree is stale'
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('保存按钮出现', () => q('[data-panel-action=save]') !== null)
        click(q('[data-panel-action=save]'))
        await waitFor(
          '显示失败原因',
          () => q('.uvf_dockFailure')?.textContent?.includes('worktree is stale') === true
        )
        expect(q('.uvf_dockFailure')?.textContent).toContain(t('dock.actionFailed'))
        expect(panelChip(), 'a rejected action keeps the changes').toBe(t('dock.unsaved'))
        expect(q('.uvf_dockFrame')?.getAttribute('src')).toBe(ZH_DEFAULT_UNIT_URL)
      })

      it('asks before closing the panel with unsaved changes', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('面板出现', () => q('.uvf_dock') !== null)
        click(q('.uvf_dockHead [aria-label="' + t('dock.close') + '"]'))
        await waitFor('出现关闭确认', () => q('[data-panel-action=close-discard]') !== null)
        expect(q('[role=alertdialog]')?.textContent).toContain(t('dock.unsavedClose'))
        expect(actionRequests, 'opening the dialog submits nothing').toHaveLength(0)
        click(q('[data-panel-action=close-cancel]'))
        await waitFor('取消后仍留在面板', () => q('[data-panel-action=close-discard]') === null)
        expect(q('.uvf_dock')).not.toBeNull()
        click(q('.uvf_dockHead [aria-label="' + t('dock.close') + '"]'))
        await waitFor('再次出现关闭确认', () => q('[data-panel-action=close-discard]') !== null)
        click(q('[data-panel-action=close-discard]'))
        await waitFor('关闭后丢弃改动', () => q('.uvf_dock') === null)
        expect(actionRequests.at(-1)).toMatchObject({ action: 'discard', worktreeId: WORKTREE })
      })

      it('warns the browser before unload while changes are unsaved', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('面板出现', () => q('.uvf_dock') !== null)
        const event = new dom.window.Event('beforeunload', { cancelable: true })
        dom.window.dispatchEvent(event)
        expect(event.defaultPrevented).toBe(true)
      })

      it('shows the saved version read-only when the branch has no unsaved changes', async () => {
        branch = branchState(false)
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('已保存状态出现', () => panelChip() === t('dock.saved'))
        expect(q('[data-panel-action=save]')?.disabled).toBe(true)
        expect(q('[data-panel-action=discard]')?.disabled).toBe(true)
        expect(q('[data-panel-action=compare]'), 'nothing to compare against').toBeNull()
      })

      it('resizes from the left splitter and remembers the width', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('调整宽度的分隔条出现', () => q('.uvf_dockSplitter') !== null)
        const splitter = q('.uvf_dockSplitter')
        pointer(splitter, 'pointerdown', { pointerId: 1, clientX: 1000 })
        await sleep(50)
        pointer(splitter, 'pointermove', { pointerId: 1, clientX: 800 })
        await waitFor('面板跟随指针变宽', () => q('.uvf_dock')?.style.width === '640px')
        await sleep(50)
        pointer(splitter, 'pointerup', { pointerId: 1, clientX: 800 })
        await waitFor(
          '松开后记住宽度',
          () => dom.window.localStorage.getItem('dsh-univer-office.panel-width') === '640'
        )
      })

      it('explains a stopped Gateway', async () => {
        gatewayPhase = 'stopped'
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor(
          'Gateway 停止时给出说明',
          () => q('.uvf_note')?.textContent?.includes(t('dock.gatewayDown')) === true
        )
        expect(q('.uvf_dockFrame')).toBeNull()
      })

      it('releases the session binding and the panel when closed', async () => {
        // Nothing unsaved, so closing is immediate: the confirmation strip only guards changes.
        branch = branchState(false)
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('面板出现', () => q('.uvf_dock') !== null)
        await waitFor('已绑定到会话', () => panelOpens.length === 1)
        click(q(`button[aria-label="${t('dock.close')}"]`))
        await waitFor('关闭后面板消失', () => q('.uvf_dock') === null)
        await waitFor('关闭时释放会话绑定', () => panelCloses === 1)
      })

      // The panel is the user's own focus: a later snapshot without a write keeps it open, and
      // only the close button (or a session change) removes it.
      it('keeps the panel open when a later snapshot writes nothing', async () => {
        renderPanel(sessionWithTargets([{ file: DEMO_FILE, worktreeId: WORKTREE }], false))
        await waitFor('面板出现', () => q('.uvf_dock') !== null)
        renderPanel(sessionWithTargets([], false))
        await sleep(300)
        expect(q('.uvf_dock')?.getAttribute('data-file')).toBe(DEMO_FILE)
      })
    })

    describe('on-demand .univer file viewer', () => {
      let viewer: { element: HTMLElement; root: any }
      let key = 0
      const renderViewer = () =>
        viewer.root.render(
          React.createElement(fileViewer.component, {
            key: 'fv' + ++key,
            ctx: fakeCtx,
            scope: { sessionId: SESSION_ID, cwd: SESSION_CWD },
            path: DEMO_FILE,
            title: 'demo.univer',
            viewerId: fileViewer.id
          })
        )
      const chips = () => viewer.element.querySelectorAll('.uvf_unit')
      const frameSrc = () =>
        viewer.element.querySelector('iframe.uvf_dockFrame')?.getAttribute('src')

      beforeAll(() => {
        viewer = mount()
      })

      beforeEach(() => {
        panelBinding = null
        panelOpens.length = 0
        gatewayPhase = 'running'
        missingFiles.delete(DEMO_FILE)
      })

      afterAll(() => {
        gatewayPhase = 'running'
        missingFiles.delete(DEMO_FILE)
        viewer.root.unmount()
        viewer.element.remove()
      })

      // Opening the file must reach the Viewer with no agent operation, and the same tab must carry
      // the review controls: there is no second surface to send the user to.
      it('opens the current branch on click with its review controls in the same tab', async () => {
        branch = branchState(true)
        renderViewer()
        await waitFor(
          '点击文件即在同一个预览页展示当前分支',
          () => frameSrc() === withLang(DEFAULT_UNIT_URL, 'zh-CN')
        )
        expect(chips()).toHaveLength(2)
        expect(viewer.element.querySelector('[data-panel-action=save]')).not.toBeNull()
        expect(viewer.element.querySelector('[data-panel-action=discard]')).not.toBeNull()
        expect(
          viewer.element.querySelector('.uvf_dock'),
          'the preview tab carries no detached panel chrome'
        ).toBeNull()
        // Switching locale re-targets the embedded Viewer instead of loading another surface.
        setActiveLocale('en')
        try {
          await waitFor(
            '语言切换后 Viewer locale 同步跟随',
            () => frameSrc() === withLang(DEFAULT_UNIT_URL, 'en-US')
          )
          expect(viewer.element.querySelectorAll('iframe')).toHaveLength(1)
        } finally {
          setActiveLocale('zh')
        }
      })

      it('shows the saved version when no branch is open, with nothing to settle', async () => {
        branch = null
        renderViewer()
        await waitFor('没有分支时展示当前版本', () => frameSrc() === trunkUrl())
        expect(chips()).toHaveLength(0)
        expect(viewer.element.querySelector('[data-panel-action=save]')?.hasAttribute('disabled')).toBe(
          true
        )
        expect(
          viewer.element.querySelector('[data-panel-action=discard]')?.hasAttribute('disabled')
        ).toBe(true)
        expect(viewer.element.querySelector('[data-panel-action=compare]')).toBeNull()
      })

      it('explains a file the Host reports missing', async () => {
        missingFiles.add(DEMO_FILE)
        renderViewer()
        await waitFor(
          '文件已不在工作区时给出说明',
          () => viewer.element.textContent?.includes(t('viewer.missing')) === true
        )
        expect(viewer.element.querySelector('iframe')).toBeNull()
      })

      // The shared state poll swallows this failure, so the phase is reported separately.
      it('offers a start action for a stopped Gateway', async () => {
        gatewayPhase = 'stopped'
        renderViewer()
        await waitFor(
          'Gateway 停止时给出启动入口',
          () => viewer.element.textContent?.includes(t('dock.gatewayDown')) === true
        )
        const start = Array.from(viewer.element.querySelectorAll('button')).find(
          (button) => button.textContent === t('dock.startGateway')
        )
        expect(start).toBeDefined()
      })
    })

    // The seat hands a body nothing but the resource address, so the address carries both the
    // session and the path; this proves the route end to end against the same fake Host API.
    it('renders the native right-Sidebar tab body from its address alone', async () => {
      const { element, root } = mount()
      try {
        panelBinding = null
        branch = null
        root.render(
          React.createElement(nativeBodyEntry.Component, {
            ctx: fakeCtx,
            useTabInfo: () => ({
              tab: {
                contentId: `dsh-resource://file/session/${SESSION_ID}/${encodeAddressPath(REL_DEMO_FILE)}`
              }
            })
          })
        )
        await waitFor(
          '原生侧边栏标签按地址解析出文件与会话并渲染 Viewer',
          () => element.querySelector('iframe.uvf_dockFrame')?.getAttribute('src') === trunkUrl()
        )
        expect(element.querySelector('[data-surface=sidebar-right]')).not.toBeNull()
        expect(element.querySelector('.uvf_unit')).toBeNull()
        expect(
          element.querySelector('[data-panel-action=save]'),
          'the same tab hosts the review controls'
        ).not.toBeNull()
      } finally {
        root.unmount()
        element.remove()
      }
    })

    // One write, one surface: an agent's work reaches the preview the user has open, and never a
    // second panel beside it. These cases drive the real registration rather than a stand-in.
    describe('single preview surface', () => {
      let target: { element: HTMLElement; root: any }
      let dock: { element: HTMLElement; root: any }
      const OTHER_UNIT_URL = `${VIEW_URL.replace(WORKTREE, OTHER_WORKTREE)}&unit=u-other0000-unit`

      beforeEach(() => {
        branch = branchState(true)
        panelBinding = null
        panelOpens.length = 0
        actionRequests.length = 0
        actionFailure = undefined
        gatewayPhase = 'running'
        missingFiles.clear()
        target = mount()
        dock = mount()
      })

      afterEach(async () => {
        target.root.unmount()
        target.element.remove()
        dock.root.unmount()
        dock.element.remove()
        sidebarNavigator = undefined
        await sleep(120)
      })

      const renderViewer = () =>
        target.root.render(
          React.createElement(fileViewer.component, {
            key: 'sp' + String(Date.now()),
            ctx: fakeCtx,
            scope: { sessionId: SESSION_ID, cwd: SESSION_CWD },
            path: DEMO_FILE,
            title: 'demo.univer',
            viewerId: fileViewer.id
          })
        )
      const renderDock = (files: any[]) =>
        dock.root.render(
          React.createElement(dockEntry.Component, dockProps(sessionWithFiles(files, false)))
        )
      const frameSrc = () =>
        target.element.querySelector('iframe.uvf_dockFrame')?.getAttribute('src')

      // The binding is the Host's report of the session's edit target, and an agent write rebinds
      // it before revealing the tab. The open preview must adopt it in place.
      it('adopts the branch an agent write binds, in the same tab', async () => {
        renderViewer()
        await waitFor('先展示当前分支', () => frameSrc() === withLang(DEFAULT_UNIT_URL, 'zh-CN'))
        branch = {
          ...branchState(true, OTHER_WORKTREE),
          worktreeUrl: VIEW_URL.replace(WORKTREE, OTHER_WORKTREE),
          units: [
            {
              unitId: 'u-other0000-unit',
              name: 'Other',
              type: 'sheet',
              kind: 'added',
              worktreeUrl: OTHER_UNIT_URL
            }
          ]
        }
        panelBinding = {
          source: null,
          file: DEMO_FILE,
          worktreeId: OTHER_WORKTREE,
          unitId: 'u-other0000-unit',
          viewerUrl: ''
        }
        await waitFor(
          '绑定变化后预览页原地跟随，不弹出第二个面',
          () => frameSrc() === withLang(OTHER_UNIT_URL, 'zh-CN')
        )
        expect(q('.uvf_dock'), 'no detached panel appears').toBeNull()
        expect(document.querySelectorAll('iframe')).toHaveLength(1)
      })

      it('accepts and rejects the change from inside the preview', async () => {
        branch = branchState(true)
        renderViewer()
        await waitFor('审阅按钮出现在预览页内', () => q('[data-panel-action=discard]') !== null)
        click(q('[data-panel-action=discard]'))
        await waitFor('确认仍在同一页内', () => q('[role=alertdialog]') !== null)
        expect(q('[role=alertdialog]')?.textContent).toContain(t('dock.confirmDiscard'))
        click(q('[data-panel-action=confirm-discard]'))
        await waitFor('丢弃后预览页回到当前版本', () => frameSrc() === trunkUrl())
        expect(actionRequests.at(-1)).toMatchObject({ action: 'discard', worktreeId: WORKTREE })
        expect(target.element.querySelector('.uvf_dock')).toBeNull()
      })

      // Nothing is rendered beside the preview, and the reveal lands on the tab's own address.
      it('reveals the open preview instead of opening a second page', async () => {
        const opened = installSidebarNavigator()
        renderViewer()
        await waitFor('预览页先展示草稿', () => frameSrc() === withLang(DEFAULT_UNIT_URL, 'zh-CN'))
        renderDock([turnFile(DEMO_FILE, WORKTREE)])
        await waitFor('写入后揭示同一个预览页', () => opened.length === 1)
        expect(opened[0]).toBe(
          `dsh-resource://file/session/${SESSION_ID}/${encodeAddressPath(DEMO_FILE)}`
        )
        expect(q('.uvf_dock'), 'the write renders no panel of its own').toBeNull()
        expect(document.querySelectorAll('iframe')).toHaveLength(1)
      })

      // A deck the user opened as `.pptx` is one tab: the agent writes its `.univer` container,
      // and the reveal must land on the deck's address instead of opening the container beside it.
      it('reveals an opened deck instead of its container', async () => {
        const opened = installSidebarNavigator()
        const deck = DEMO_FILE.replace(/\.univer$/, '.pptx')
        branch = branchState(true)
        target.root.render(
          React.createElement(fileViewer.component, {
            key: 'deck',
            ctx: fakeCtx,
            scope: { sessionId: SESSION_ID, cwd: SESSION_CWD },
            path: deck,
            title: 'demo.pptx',
            viewerId: fileViewer.id
          })
        )
        await waitFor(
          '演示文稿导入后展示草稿',
          () => frameSrc() === withLang(DEFAULT_UNIT_URL, 'zh-CN')
        )
        renderDock([turnFile(DEMO_FILE, WORKTREE)])
        await waitFor('写入后揭示演示文稿那一个标签页', () => opened.length === 1)
        expect(opened[0]).toBe(
          `dsh-resource://file/session/${SESSION_ID}/${encodeAddressPath(deck)}`
        )
      })
    })

    it('releases every registration on unload', () => {
      for (const dispose of effects.toReversed()) if (typeof dispose === 'function') dispose()
      expect(localeListeners.size).toBe(0)
    })

    // Real Cordis exercises dependency visibility and optional child fibers. Set
    // UNIVER_DSH_RUNTIME_ROOT to a separately installed DSH package directory to run the same
    // lifecycle against that release's Cordis.
    it('activates and tears down under real Cordis', async () => {
      const runtimeRequire = process.env.UNIVER_DSH_RUNTIME_ROOT
        ? createRequire(join(process.env.UNIVER_DSH_RUNTIME_ROOT, 'package.json'))
        : repoRequire
      const { Context } = await import(
        pathToFileURL(runtimeRequire.resolve('@deepseek-ai/cordis')).href
      )
      const context = new Context()
      const liveEntries = new Map<any, string>()
      try {
        context.provide('slots', {
          inject(_name: string, callback: () => unknown) {
            return context.effect(callback)
          },
          register(options: any) {
            liveEntries.set(options, options.name)
            return () => liveEntries.delete(options)
          }
        })
        context.provide('locale', fakeCtx.locale)
        context.provide('conversation', {})
        const fiber = context.plugin(pluginExports)
        await fiber
        expect(fiber.state, 'Client waits for uiConversation before activating').toBe(0)
        context.provide('uiConversation', { events: conversationEventRegistry })
        await fiber
        expect(fiber.state, 'absent optional services do not block activation').toBe(2)
        const names = [...liveEntries.values()]
        expect(names).toContain('conversation.input.dock')
        expect(names, 'the chat log never hosts a Univer card').not.toContain(
          'conversation.chat.turnTail'
        )
        await fiber.dispose()
      } finally {
        await context.fiber.dispose()
      }
      expect(liveEntries.size, 'teardown leaves no slot registrations').toBe(0)
    })
  })
}
