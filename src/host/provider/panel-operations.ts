import { stat } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import type { OpenPanel, PanelImageSelection, PanelOpenResult } from '../../shared/wire/panel.ts'
import { VIEWER_BASE } from '../../shared/wire/viewer.ts'
import { GatewayClient, gatewayErrorMessage } from '../adapters/gateway/client.ts'
import { GatewayFileApi, fileKeyOf } from '../adapters/gateway/file-api.ts'
import { GatewayWorktreeApi } from '../adapters/gateway/worktree-api.ts'
import { isRecord, mapUnits, mapWorktrees, type GatewayUnit } from '../adapters/gateway/mapping.ts'
import { UniverError } from '../service/errors.ts'
import type {
  BindPanelRequest,
  JsonValue,
  OpenPanelSourceRequest,
  UniverUnitKind
} from '../service/types.ts'
import { presentationChanged, presentationSiblingExists } from './presentation-files.ts'

/** Imports one Office source into a JSON Unit snapshot. */
export interface PanelSourceImporter {
  import(
    sourcePath: string,
    docType: undefined,
    signal?: AbortSignal
  ): Promise<{ readonly kind: UniverUnitKind; readonly snapshot: JsonValue }>
}

const SLIDE_UNIT_TYPE = 3

/**
 * Owns the right panel's browser focus and the import that backs it.
 *
 * Opening a local deck is user intent rather than an agent write, yet the content write still goes
 * through an explicit `draft` worktree: the Host creates the container, imports the source once
 * into a draft dedicated to that import, and keeps the resulting binding in memory. That binding
 * is browser state — the panel is gone after a reload — so it deliberately survives no restart,
 * while every fact the model can observe about it arrives through `univer_status`.
 *
 * Binding accepts any Unit kind: the right panel is the only Univer surface in the browser, so a
 * Sheet, Doc, Base, or Board target must be projected exactly like a Slide one.
 */
export class PanelOperations {
  private readonly bindings = new Map<string, OpenPanel>()
  /**
   * The presentation each container caches, for this Host run.
   *
   * It outlives a panel binding: a tab closed and reopened, or a branch the Host recreates after a
   * save, must keep reporting the deck rather than the container the agent writes. Entries are
   * absolute authorized paths, so they are already canonical and safe as keys.
   */
  private readonly sources = new Map<string, string>()

  constructor(
    private readonly gatewayRequestTimeoutMs: number,
    private readonly gatewayMutationTimeoutMs: number,
    private readonly importer: PanelSourceImporter,
    /**
     * When this Host process started. A branch older than this belongs to a run the user already
     * left, so opening the file again starts from the last saved version instead of restoring work
     * that was never saved.
     */
    private readonly startedAt: number
  ) {}

  /** The presentation one container caches, when this Host run knows one. */
  presentationOf(file: string): Promise<string | undefined> {
    const known = this.sources.get(file)
    if (known !== undefined) return Promise.resolve(known)
    // A container created beside a deck is that deck's cache even after a Host restart; the sibling
    // is what the import named it after, so no sidecar has to record the provenance.
    return presentationSiblingExists(file).then((source) => {
      if (source !== undefined) this.sources.set(file, source)
      return source
    })
  }

  /** Import one local deck into the editable branch and bind it to the session. */
  async open(
    gateway: string,
    request: OpenPanelSourceRequest,
    signal?: AbortSignal
  ): Promise<PanelOpenResult> {
    signal?.throwIfAborted()
    const client = new GatewayClient(gateway, this.gatewayMutationTimeoutMs)
    await this.ensureContainer(client, request.file)
    this.sources.set(request.file, request.source)
    await this.discardAbandoned(client, request.file, signal)
    // A deck regenerated outside the plugin is authoritative: the cache is rebuilt before the panel
    // shows anything, so the preview always matches the file on disk.
    if (await presentationChanged(request.source, request.file)) {
      await this.reimport(client, request.file, request.source, signal)
    }
    const found = await this.findImportDraft(client, request.file)
    const worktreeId = found?.worktreeId ?? (await this.createDraft(client, request.file))
    let unitId = found?.unit.unitId
    if (unitId === undefined) {
      unitId = await this.adoptOrImport(client, request, worktreeId, signal)
    }
    return this.commit(request.sessionId, {
      source: request.source,
      file: request.file,
      worktreeId,
      unitId,
      viewerUrl: viewerUrlOf(request.file, worktreeId, unitId)
    })
  }

  /**
   * Bind an already-authorized Unit to the session, materializing the editable branch if needed.
   *
   * Opening a `.univer` is the user's way into its one branch, so this never fails for a file that
   * has content: an existing branch is adopted, and a file with none gains one holding the saved
   * version. Only a container without any Unit has nothing to show.
   */
  async bind(
    gateway: string,
    request: BindPanelRequest,
    signal?: AbortSignal
  ): Promise<PanelOpenResult> {
    signal?.throwIfAborted()
    const client = new GatewayClient(gateway, this.gatewayRequestTimeoutMs)
    await this.discardAbandoned(client, request.file, signal)
    // A reveal driven by an agent write names the container, not the deck the user opened. Keeping
    // the source on the binding is what makes the tab stay the deck and lets Save write back to it.
    const source = request.source ?? (await this.presentationOf(request.file)) ?? null
    const worktreeId = request.worktreeId ?? (await this.ensureBranch(client, request.file))
    const units = mapUnits(await new GatewayWorktreeApi(client).listUnits(request.file, worktreeId))
    const unit = request.unitId === undefined ? defaultUnit(units) : findUnit(units, request.unitId)
    if (unit === undefined)
      throw new UniverError('The selected worktree holds no matching Unit.', 'UNIT_NOT_FOUND')
    return this.commit(request.sessionId, {
      source,
      file: request.file,
      worktreeId,
      unitId: unit.unitId,
      viewerUrl: viewerUrlOf(request.file, worktreeId, unit.unitId)
    })
  }

  /** Release the session's binding; a panel that closed must not keep routing the conversation. */
  release(sessionId: string): void {
    this.bindings.delete(sessionId)
  }

  /** The panel target one session currently holds, if any. */
  of(sessionId: string): OpenPanel | undefined {
    return this.bindings.get(sessionId)
  }

  /** Drop every binding on unload. */
  clear(): void {
    this.bindings.clear()
  }

  /**
   * Replace the image the session's Viewer has selected.
   *
   * A report is browser state that can arrive after the panel closed or moved on, so a report only
   * lands on a matching binding and everything else is dropped silently.
   */
  select(sessionId: string, selection: PanelImageSelection | null): void {
    const panel = this.bindings.get(sessionId)
    if (panel === undefined) return
    if (sameSelection(panel.selection, selection)) return
    this.bindings.set(sessionId, { ...panel, selection })
  }

  private commit(sessionId: string, panel: Omit<OpenPanel, 'selection'>): PanelOpenResult {
    const bound: OpenPanel = { ...panel, selection: null }
    this.bindings.set(sessionId, bound)
    return { ok: true, panel: bound }
  }

  /**
   * Discard branches left behind by a previous Host run.
   *
   * Unsaved work is deliberately not persisted: a user who quits DSH or the browser expects the
   * file to reopen at the last saved version, so any branch this process did not create is stale.
   * A branch that a concurrent writer already merged or discarded fails here and is simply no
   * longer active, which is the same outcome.
   */
  private async discardAbandoned(
    client: GatewayClient,
    file: string,
    signal?: AbortSignal
  ): Promise<void> {
    const api = new GatewayWorktreeApi(client)
    const abandoned = mapWorktrees(await new GatewayFileApi(client).listWorktrees(file)).filter(
      (worktree) =>
        (worktree.status === 'draft' || worktree.status === 'ready') &&
        abandonedBy(worktree.createdAt, this.startedAt)
    )
    for (const worktree of abandoned) {
      signal?.throwIfAborted()
      try {
        await api.action(file, worktree.worktreeId, 'discard')
      } catch {
        // Expected when a concurrent writer settled the same branch first; it is gone either way.
      }
    }
  }

  /** Adopt the file's active branch, or create the one that holds the saved version. */
  private async ensureBranch(client: GatewayClient, file: string): Promise<string> {
    const active = mapWorktrees(await new GatewayFileApi(client).listWorktrees(file))
      .filter((worktree) => worktree.status === 'draft' || worktree.status === 'ready')
      .sort((left, right) => (right.createdAt ?? '').localeCompare(left.createdAt ?? ''))[0]
    if (active === undefined) return this.createDraft(client, file)
    if (active.status === 'ready') {
      requireGatewaySuccess(
        await new GatewayWorktreeApi(client).action(file, active.worktreeId, 'reopen'),
        'Gateway rejected reopening the edit branch.'
      )
    }
    return active.worktreeId
  }

  /** Create the `.univer` container once; the Gateway materializes it at the authorized path. */
  private async ensureContainer(client: GatewayClient, file: string): Promise<void> {
    if (await exists(file)) return
    requireGatewaySuccess(
      await new GatewayFileApi(client).create(file),
      'Gateway rejected the Univer file creation.'
    )
    if (!(await exists(file)))
      throw new UniverError(
        'The Gateway did not materialize the Univer container.',
        'GATEWAY_INVALID_RESPONSE'
      )
  }

  /**
   * Find the newest draft worktree that already holds an imported Unit.
   *
   * Imports live in a draft of their own so an unrelated sheet or document draft is never reused,
   * which also makes a repeated open of the same deck idempotent.
   */
  private async findImportDraft(
    client: GatewayClient,
    file: string
  ): Promise<{ readonly worktreeId: string; readonly unit: GatewayUnit } | undefined> {
    const drafts = mapWorktrees(await new GatewayFileApi(client).listWorktrees(file))
      .filter((worktree) => worktree.status === 'draft')
      .sort((left, right) => (right.createdAt ?? '').localeCompare(left.createdAt ?? ''))
    const api = new GatewayWorktreeApi(client)
    for (const draft of drafts) {
      const unit = defaultUnit(mapUnits(await api.listUnits(file, draft.worktreeId)))
      if (unit !== undefined) return { worktreeId: draft.worktreeId, unit }
    }
    return undefined
  }

  private async createDraft(
    client: GatewayClient,
    file: string,
    name = 'Univer panel import'
  ): Promise<string> {
    const created = await new GatewayWorktreeApi(client).create(file, name)
    requireGatewaySuccess(created, 'Gateway rejected the worktree creation.')
    const worktreeId = recordString(created, 'worktreeId')
    if (worktreeId === undefined)
      throw new UniverError(
        'Gateway worktree result is missing worktreeId.',
        'GATEWAY_INVALID_RESPONSE'
      )
    return worktreeId
  }

  /**
   * Take the deck a fresh branch already holds, or import it when the trunk has none.
   *
   * A branch is seeded from the saved version, so after a save or a re-import the deck is already
   * there. Importing again would add a second Slide Unit to the same container.
   */
  private async adoptOrImport(
    client: GatewayClient,
    request: OpenPanelSourceRequest,
    worktreeId: string,
    signal?: AbortSignal
  ): Promise<string> {
    const existing = defaultUnit(
      mapUnits(await new GatewayWorktreeApi(client).listUnits(request.file, worktreeId))
    )
    if (existing !== undefined) return existing.unitId
    return this.importUnit(client, request, worktreeId, signal)
  }

  /**
   * Rebuild the container from a deck that changed on disk.
   *
   * The container is only a cache of the presentation, so a newer source replaces it wholesale: one
   * import, one branch that adds the new Unit and retires the old one, then a merge, which leaves the
   * trunk with exactly the deck the file now holds. Unsaved work in any branch is based on the old
   * deck and is dropped with those branches.
   */
  private async reimport(
    client: GatewayClient,
    file: string,
    source: string,
    signal?: AbortSignal
  ): Promise<void> {
    const api = new GatewayWorktreeApi(client)
    const active = mapWorktrees(await new GatewayFileApi(client).listWorktrees(file)).filter(
      (worktree) => worktree.status === 'draft' || worktree.status === 'ready'
    )
    for (const worktree of active) {
      signal?.throwIfAborted()
      try {
        await api.action(file, worktree.worktreeId, 'discard')
      } catch {
        // A concurrent writer settled the same branch first; it is gone either way.
      }
    }
    const imported = await this.importer.import(source, undefined, signal)
    if (imported.kind !== 'slide')
      throw new UniverError(
        'Only a presentation source can be imported into the Univer panel.',
        'IMPORT_FORMAT_UNSUPPORTED'
      )
    const draft = await this.createDraft(client, file, 'Presentation re-import')
    try {
      const before = new Set(mapUnits(await api.listUnits(file, draft)).map((unit) => unit.unitId))
      requireGatewaySuccess(
        await api.createUnit(file, draft, 'slide', deckName(source), imported.snapshot),
        'Gateway rejected the re-imported Unit.'
      )
      const units = mapUnits(await api.listUnits(file, draft))
      const created = units.find(
        (unit) => unit.type === SLIDE_UNIT_TYPE && !before.has(unit.unitId)
      )
      if (created === undefined)
        throw new UniverError(
          'Gateway did not report the re-imported Unit.',
          'GATEWAY_INVALID_RESPONSE'
        )
      // The retired deck is a trunk Unit in this branch, so the merge removes it from the trunk.
      for (const unit of units) {
        signal?.throwIfAborted()
        if (unit.unitId === created.unitId || unit.type !== SLIDE_UNIT_TYPE) continue
        requireGatewaySuccess(
          await api.removeUnit(file, draft, unit.unitId),
          'Gateway rejected removing the replaced presentation Unit.'
        )
      }
      requireGatewaySuccess(
        await api.action(file, draft, 'merge'),
        'Gateway rejected merging the re-imported presentation.'
      )
    } catch (error) {
      // Leaving a half-built branch behind would make the next open reuse a partial import.
      await api.action(file, draft, 'discard').catch(() => undefined)
      throw error
    }
  }

  private async importUnit(
    client: GatewayClient,
    request: OpenPanelSourceRequest,
    worktreeId: string,
    signal?: AbortSignal
  ): Promise<string> {
    const imported = await this.importer.import(request.source, undefined, signal)
    if (imported.kind !== 'slide')
      throw new UniverError(
        'Only a presentation source can be imported into the Univer panel.',
        'IMPORT_FORMAT_UNSUPPORTED'
      )
    const api = new GatewayWorktreeApi(client)
    const before = new Set(
      mapUnits(await api.listUnits(request.file, worktreeId)).map((unit) => unit.unitId)
    )
    const created = await api.createUnit(
      request.file,
      worktreeId,
      'slide',
      deckName(request.source),
      imported.snapshot
    )
    requireGatewaySuccess(created, 'Gateway rejected the imported Unit.')
    const added = mapUnits(await api.listUnits(request.file, worktreeId)).find(
      (unit) => unit.type === SLIDE_UNIT_TYPE && !before.has(unit.unitId)
    )
    const reported = recordString(created, 'unitId')
    if (reported === undefined && added === undefined)
      throw new UniverError('Gateway did not report the imported Unit.', 'GATEWAY_INVALID_RESPONSE')
    return added?.unitId ?? reported!
  }
}

/** A presentation Unit is the expected target when the caller names none. */
function defaultUnit(units: readonly GatewayUnit[]): GatewayUnit | undefined {
  return units.find((unit) => unit.type === SLIDE_UNIT_TYPE) ?? units[0]
}

function findUnit(units: readonly GatewayUnit[], unitId: string): GatewayUnit | undefined {
  return units.find((unit) => unit.unitId === unitId)
}

function deckName(source: string): string {
  const name = basename(source)
  const extension = extname(name)
  return extension.length === 0 ? name : name.slice(0, -extension.length)
}

function viewerUrlOf(file: string, worktreeId: string, unitId: string): string {
  const base = `${VIEWER_BASE}/?file=${encodeURIComponent(fileKeyOf(file))}`
  return (
    `${base}&worktree=${encodeURIComponent(worktreeId)}&mode=embedded&scope=worktree` +
    `&unit=${encodeURIComponent(unitId)}`
  )
}

async function exists(file: string): Promise<boolean> {
  try {
    await stat(file)
    return true
  } catch (error) {
    // Only a missing container means "create it"; a permission failure must surface.
    if (nodeErrorCode(error) === 'ENOENT') return false
    throw error
  }
}

function nodeErrorCode(error: unknown): string | undefined {
  if (typeof error !== 'object' || error === null || !('code' in error)) return undefined
  return typeof error.code === 'string' ? error.code : undefined
}

function requireGatewaySuccess(value: JsonValue, fallback: string): void {
  if (isRecord(value) && isRecord(value.error) && value.error.code === 0)
    throw new UniverError(gatewayErrorMessage(value) ?? fallback, 'GATEWAY_REQUEST_REJECTED')
  if (isRecord(value) && value.ok === false)
    throw new UniverError(gatewayErrorMessage(value) ?? fallback, 'GATEWAY_REQUEST_REJECTED')
}

function recordString(value: JsonValue, key: string): string | undefined {
  if (!isRecord(value)) return undefined
  const found = value[key]
  return typeof found === 'string' && found.length > 0 ? found : undefined
}

/** True when a branch was created before this Host process, so it is abandoned unsaved work. */
function abandonedBy(createdAt: string | undefined, startedAt: number): boolean {
  if (createdAt === undefined) return true
  const created = Date.parse(createdAt)
  // An unparseable timestamp is treated as current: dropping user work needs positive evidence.
  return Number.isFinite(created) && created < startedAt
}

/** Selection identity compared field by field; a repeated report must not churn the binding. */
function sameSelection(
  left: PanelImageSelection | null,
  right: PanelImageSelection | null
): boolean {
  if (left === null || right === null) return left === right
  return (
    left.unitId === right.unitId &&
    left.pageId === right.pageId &&
    left.elementId === right.elementId
  )
}
