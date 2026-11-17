import type { Context } from '@deepseek-ai/cordis'
import { createStandardApiReference, type ApiReference } from '@univer-cli/api-reference'
import type { ResolvedConfig } from '../config.ts'
import type { WorktreeActionResult } from '../../shared/wire/actions.ts'
import type { OpenPanel, PanelOpenResult } from '../../shared/wire/panel.ts'
import type { FileBranch, FileState, ChangedUnit } from '../../shared/wire/state.ts'
import type { EnsureGatewayResult, GatewayStatus } from '../../shared/wire/status.ts'
import { VIEWER_BASE } from '../../shared/wire/viewer.ts'
import { GatewayClient, gatewayErrorMessage } from '../adapters/gateway/client.ts'
import { GatewayFileApi, fileKeyOf } from '../adapters/gateway/file-api.ts'
import { GatewayWorktreeApi } from '../adapters/gateway/worktree-api.ts'
import { isRecord, mapUnits, mapWorktrees, unitKind } from '../adapters/gateway/mapping.ts'
import type {
  ApiReferenceRequest,
  BindPanelRequest,
  ExecuteUnitContentRequest,
  ExportUnitContentRequest,
  FileStateRequest,
  FileStatusRequest,
  ImportUnitContentRequest,
  InspectUnitContentRequest,
  LintUnitLayoutRequest,
  OpenPanelSourceRequest,
  PanelSelectionRequest,
  PrintPdfUnitRequest,
  ResourceOperationRequest,
  ScreenshotServiceResult,
  ScreenshotUnitRequest,
  CompileSvgRequest,
  EditSlideImageRequest,
  JsonValue,
  NewUniverFileRequest,
  UnitOperationRequest,
  UniverApiResult,
  UniverOperationResult,
  UniverResourceResult,
  UniverUnitKind,
  WorktreeActionRequest,
  WorktreeOperationRequest
} from '../service/types.ts'
import { UniverError } from '../service/errors.ts'
import {
  unitId as brandedUnitId,
  worktreeId as brandedWorktreeId,
  type UniverFilePath,
  type WorkspacePath,
  type WorktreeId
} from '../service/identifiers.ts'
import { UniverService } from '../service/univer-service.ts'
import { assertAuthorizedPath } from '../service/workspace.ts'
import { GatewaySupervisor } from '../processes/gateway/supervisor.ts'
import { UnitContentOperations } from './unit-content-operations.ts'
import { generateImage } from './image-generation-client.ts'
import {
  imageDataUrl,
  locateSlideImage,
  readSlideImageBytes,
  replaceSlideImageCode
} from './slide-image.ts'
import { PanelOperations } from './panel-operations.ts'
import { StateCache } from './state-cache.ts'
import { WorktreeOperations } from './worktree-operations.ts'
import { RenderOperations } from './render-operations.ts'
import { RenderSourceOperations } from './render-source-operations.ts'
import { ResourceOperations } from './resource-operations.ts'
import {
  commitStagedPresentation,
  discardStagedPresentation,
  markPresentationImported,
  stagedPresentationPath
} from './presentation-files.ts'
import { keepsUserEdits, PresentationEpisodes } from './presentation-episodes.ts'

/** Gateway Unit type for a Presentation. */
const SLIDE_UNIT_TYPE = 3

/** Local Service Provider backed by the bundled Gateway and Unit content worker. */
export class GatewayUniverService extends UniverService {
  private readonly gatewaySupervisor: GatewaySupervisor
  private readonly unitContent: UnitContentOperations
  private readonly worktrees: WorktreeOperations
  private readonly render: RenderOperations
  private readonly renderSources: RenderSourceOperations
  private readonly resourceOperations: ResourceOperations
  private readonly panels: PanelOperations
  private readonly episodes = new PresentationEpisodes()
  private readonly stateCache: StateCache<string, FileState>
  private readonly unitCache: StateCache<string, readonly ChangedUnit[]>
  private readonly api: ApiReference

  constructor(
    ctx: Context,
    private readonly config: ResolvedConfig
  ) {
    super(ctx)
    this.gatewaySupervisor = new GatewaySupervisor(config)
    this.unitContent = new UnitContentOperations(
      config.gatewayRequestTimeoutMs,
      config.unitContentCommitTimeoutMs,
      config.unitContentOperationTimeoutMs,
      {
        maxEntries: config.unitContentRuntimeMaxEntries,
        idleTtlMs: config.unitContentRuntimeIdleTtlMs
      }
    )
    this.worktrees = new WorktreeOperations(
      config.gatewayRequestTimeoutMs,
      config.gatewayMutationTimeoutMs
    )
    this.render = new RenderOperations(config.browserExecutablePath)
    this.renderSources = new RenderSourceOperations(
      this.unitContent,
      config.gatewayRequestTimeoutMs
    )
    this.resourceOperations = new ResourceOperations(
      config.resourceCacheRoot,
      config.resourceDownloadTimeoutMs
    )
    this.panels = new PanelOperations(
      config.gatewayRequestTimeoutMs,
      config.gatewayMutationTimeoutMs,
      this.unitContent,
      Date.now()
    )
    this.stateCache = new StateCache(config.stateCacheTtlMs)
    this.unitCache = new StateCache(config.unitCacheTtlMs)
    this.api = createStandardApiReference()
    ctx.effect(() => async () => this.dispose(), 'univer: Gateway supervisor')
  }

  /** Current Gateway status. */
  gatewayStatus(): Promise<GatewayStatus> {
    return this.gatewaySupervisor.status()
  }

  /** Ensure the bundled Gateway is available. */
  ensureGateway(): Promise<EnsureGatewayResult> {
    return this.gatewaySupervisor.ensure()
  }

  /** Return cached collaboration state for one authorized file. */
  async fileState(request: FileStateRequest): Promise<FileState> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    return this.stateCache.get(request.file, () => this.computeFileState(request.file))
  }

  /** Apply a browser review decision and return the refreshed state. */
  async worktreeAction(request: WorktreeActionRequest): Promise<WorktreeActionResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    const available = await this.ensureGateway()
    if (!available.ok) return { ok: false, reason: available.reason }
    try {
      const target = {
        workspace: request.workspace,
        file: request.file,
        worktreeId: request.worktreeId
      }
      if (request.action === 'merge') {
        await this.settlePresentation(available.gateway, target)
      } else {
        await this.discardPresentation(available.gateway, target)
      }
      this.invalidate(request.file, request.worktreeId)
      return {
        ok: true,
        action: request.action,
        worktreeId: request.worktreeId,
        state: await this.fileState({ workspace: request.workspace, file: request.file })
      }
    } catch (error) {
      this.invalidate(request.file, request.worktreeId)
      const state = await this.fileState({
        workspace: request.workspace,
        file: request.file
      }).catch(() => undefined)
      return {
        ok: false,
        reason: error instanceof Error ? error.message : String(error),
        ...(state === undefined ? {} : { state })
      }
    }
  }

  /** Create one empty Univer container without an implicit Unit. */
  async newFile(
    request: NewUniverFileRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, false)
    signal?.throwIfAborted()
    const gateway = await this.requireGateway()
    const result = await new GatewayFileApi(
      new GatewayClient(gateway, this.config.gatewayMutationTimeoutMs)
    ).create(request.file)
    requireGatewaySuccess(result, 'Gateway rejected the Univer file creation.')
    signal?.throwIfAborted()
    this.stateCache.delete(request.file)
    return {
      ok: true,
      operation: 'new',
      file: request.file,
      result: { filePath: request.file, created: true }
    }
  }

  /** Return trunk and worktree identities needed before model edits. */
  async status(request: FileStatusRequest, signal?: AbortSignal): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    signal?.throwIfAborted()
    const gateway = await this.requireGateway()
    const client = new GatewayClient(gateway, this.config.gatewayRequestTimeoutMs)
    const fileApi = new GatewayFileApi(client)
    const worktreeApi = new GatewayWorktreeApi(client)
    const [trunkValue, worktreeValue] = await Promise.all([
      fileApi.listUnits(request.file),
      fileApi.listWorktrees(request.file)
    ])
    const worktrees = mapWorktrees(worktreeValue)
    let trunkUnits = mapUnits(trunkValue)
    if (request.unitId !== undefined)
      trunkUnits = trunkUnits.filter((unit) => unit.unitId === request.unitId)
    // Edits live in the file's single branch, so that is the only one the model is shown. An
    // explicit id still works for a caller that kept one from an earlier Turn.
    const branch =
      request.worktreeId === undefined
        ? newestActive(worktrees)
        : worktrees.find((worktree) => worktree.worktreeId === request.worktreeId)
    if (request.worktreeId !== undefined && branch === undefined) {
      throw new UniverError(`Worktree ${request.worktreeId} was not found.`, 'WORKTREE_NOT_FOUND')
    }
    const branchUnits =
      branch === undefined
        ? undefined
        : mapUnits(await worktreeApi.listUnits(request.file, branch.worktreeId)).filter(
            (unit) => request.unitId === undefined || unit.unitId === request.unitId
          )
    signal?.throwIfAborted()
    return {
      ok: true,
      operation: 'status',
      file: request.file,
      result: {
        trunk: { units: trunkUnits.map(unitResult) },
        ...(branch === undefined || branchUnits === undefined
          ? {}
          : {
              branch: {
                ...worktreeResult(branch),
                units: branchUnits.map(unitResult)
              }
            })
      }
    }
  }

  /** Ensure and return the file's one editable branch. */
  async worktree(
    request: WorktreeOperationRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    signal?.throwIfAborted()
    const gateway = await this.requireGateway()
    if (request.action === 'create') {
      const worktreeId = await this.ensureEditBranch(gateway, request.file, request.name)
      signal?.throwIfAborted()
      return {
        ok: true,
        operation: 'worktree',
        file: request.file,
        result: { action: 'create', worktreeId, status: 'draft' }
      }
    }
    // An agent-driven save or discard settles the presentation exactly like the toolbar does, so the
    // original deck and the pre-agent state stay the one definition of those two decisions.
    const target = {
      workspace: request.workspace,
      file: request.file,
      worktreeId: request.worktreeId
    }
    if (request.action === 'merge') await this.settlePresentation(gateway, target)
    else await this.discardPresentation(gateway, target)
    signal?.throwIfAborted()
    this.invalidate(request.file, request.worktreeId)
    return {
      ok: true,
      operation: 'worktree',
      file: request.file,
      result: { action: request.action, worktreeId: request.worktreeId, status: 'settled' }
    }
  }

  /** Create or remove one Unit inside the editable branch. */
  async unit(request: UnitOperationRequest, signal?: AbortSignal): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    signal?.throwIfAborted()
    const gateway = await this.requireGateway()
    const worktreeId = await this.resolveEditBranch(gateway, request.file, request.worktreeId)
    if (request.action === 'remove') {
      await this.captureBaseline(gateway, request.file, worktreeId, request.unitId, signal)
    }
    const api = new GatewayWorktreeApi(
      new GatewayClient(gateway, this.config.gatewayMutationTimeoutMs)
    )
    const result =
      request.action === 'create'
        ? await api.createUnit(request.file, worktreeId, request.kind, request.name)
        : await api.removeUnit(request.file, worktreeId, request.unitId)
    requireGatewaySuccess(result, `Gateway rejected Unit ${request.action}.`)
    signal?.throwIfAborted()
    this.invalidate(request.file, worktreeId)
    return {
      ok: true,
      operation: 'unit',
      file: request.file,
      result: { action: request.action, worktreeId, ...asRecord(result) }
    }
  }

  /** Inspect one explicit Unit in trunk or a worktree. */
  async inspectUnitContent(
    request: InspectUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    const gateway = await this.requireGateway()
    return this.unitContent.inspect(gateway, request, signal)
  }

  /** Execute Facade code against one explicit Unit in the editable branch. */
  async executeUnitContent(
    request: ExecuteUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    const gateway = await this.requireGateway()
    const worktreeId = await this.resolveEditBranch(gateway, request.file, request.worktreeId)
    await this.captureBaseline(gateway, request.file, worktreeId, request.unitId, signal)
    const result = await this.unitContent.execute(
      gateway,
      request.file,
      request.code,
      worktreeId,
      request.unitId,
      signal
    )
    this.invalidate(request.file, worktreeId)
    return result
  }

  /** Import one Office file as a new Unit in the editable branch. */
  async importUnitContent(
    request: ImportUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await Promise.all([
      assertAuthorizedPath(request.workspace, request.file, true),
      assertAuthorizedPath(request.sourceWorkspace, request.source, true)
    ])
    const gateway = await this.requireGateway()
    const worktreeId = await this.resolveEditBranch(gateway, request.file, request.worktreeId)
    const imported = await this.unitContent.import(request.source, request.docType, signal)
    const result = await new GatewayWorktreeApi(
      new GatewayClient(gateway, this.config.gatewayMutationTimeoutMs)
    ).createUnit(request.file, worktreeId, imported.kind, request.name, imported.snapshot)
    requireGatewaySuccess(result, 'Gateway rejected the imported Unit.')
    signal?.throwIfAborted()
    this.invalidate(request.file, worktreeId)
    return {
      ok: true,
      operation: 'import',
      file: request.file,
      result: { sourcePath: request.source, worktreeId, ...asRecord(result) }
    }
  }

  /** Export one explicit Unit from trunk or a worktree. */
  async exportUnitContent(
    request: ExportUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await Promise.all([
      assertAuthorizedPath(request.workspace, request.file, true),
      assertAuthorizedPath(request.outputWorkspace, request.output, false)
    ])
    const gateway = await this.requireGateway()
    return this.unitContent.export(gateway, request, signal)
  }

  /** Analyze deterministic Slide layout facts without producing screenshots. */
  async lintUnitLayout(
    request: LintUnitLayoutRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    const gateway = await this.requireGateway()
    const source = await this.unitContent.renderSource(
      gateway,
      request.file,
      request.unitId,
      request.worktreeId,
      signal
    )
    const result = await this.render.lint(source, request.pages, signal)
    return { ok: true, operation: 'lint', file: request.file, result }
  }

  /** Render one explicit Unit into workspace PNGs for model visual verification. */
  async screenshotUnit(
    request: ScreenshotUnitRequest,
    signal?: AbortSignal
  ): Promise<ScreenshotServiceResult> {
    await Promise.all([
      assertAuthorizedPath(request.workspace, request.file, true),
      assertAuthorizedPath(request.outputWorkspace, request.output, false)
    ])
    const gateway = await this.requireGateway()
    const source = await this.renderSources.load(
      gateway,
      request.file,
      request.unitId,
      request.worktreeId,
      signal
    )
    const result = await this.render.screenshot({
      source,
      output: request.output,
      maxPages: this.config.screenshotMaxPages,
      maxPixels: this.config.screenshotMaxPixels,
      ...(request.target === undefined ? {} : { target: request.target }),
      ...(signal === undefined ? {} : { signal })
    })
    await Promise.all(
      result.images.map(async (image) => {
        await assertAuthorizedPath(request.outputWorkspace, image.path, true)
      })
    )
    return { ok: true, operation: 'screenshot', file: request.file, result }
  }

  /** Print one explicit Unit from trunk or a worktree to PDF. */
  async printUnitPdf(
    request: PrintPdfUnitRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await Promise.all([
      assertAuthorizedPath(request.workspace, request.file, true),
      assertAuthorizedPath(request.outputWorkspace, request.output, false)
    ])
    const gateway = await this.requireGateway()
    const source = await this.renderSources.load(
      gateway,
      request.file,
      request.unitId,
      request.worktreeId,
      signal
    )
    const result = await this.render.printPdf({
      source,
      output: request.output,
      ...(signal === undefined ? {} : { signal })
    })
    await assertAuthorizedPath(request.outputWorkspace, result.output, true)
    return { ok: true, operation: 'print-pdf', file: request.file, result }
  }

  /** Compile one SVG and commit the generated Slide mutations to the editable branch. */
  async compileSvg(
    request: CompileSvgRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await Promise.all([
      assertAuthorizedPath(request.workspace, request.file, true),
      assertAuthorizedPath(request.sourceWorkspace, request.source, true)
    ])
    if (!Number.isSafeInteger(request.page) || request.page < 1) {
      throw new UniverError('SVG target page must be a positive integer.', 'INVALID_REQUEST')
    }
    const gateway = await this.requireGateway()
    const worktreeId = await this.resolveEditBranch(gateway, request.file, request.worktreeId)
    await this.captureBaseline(gateway, request.file, worktreeId, request.unitId, signal)
    const compiled = await this.render.compileSvg({
      source: request.source,
      workspace: request.sourceWorkspace,
      page: request.page,
      ...(request.mode === undefined ? {} : { mode: request.mode }),
      ...(signal === undefined ? {} : { signal })
    })
    const execution = await this.unitContent.execute(
      gateway,
      request.file,
      compiled.code,
      worktreeId,
      request.unitId,
      signal
    )
    this.invalidate(request.file, worktreeId)
    return {
      ok: true,
      operation: 'compile-svg',
      file: request.file,
      result: {
        sourcePath: request.source,
        page: compiled.page,
        mode: compiled.mode,
        viewport: compiled.viewport,
        warnings: [...compiled.warnings],
        lints: [...compiled.lints],
        textMeasure: compiled.textMeasure,
        execution: execution.result
      }
    }
  }

  /**
   * Replace one selected Slide image with a model-generated one.
   *
   * The reference bytes come from the Unit itself rather than the browser, so the model edits what
   * the branch actually holds. The result lands as a normal Facade mutation in the edit branch, which
   * keeps Save/Discard and the collaboration refresh exactly as they are for any other edit.
   */
  async editSlideImage(
    request: EditSlideImageRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    const apiKey = this.config.readImageApiKey()
    if (apiKey === undefined) {
      throw new UniverError(
        'Image editing needs a Volcengine Ark API key. Enter it in the dsh-univer-office plugin details, or set imageApiKey or ARK_API_KEY.',
        'IMAGE_API_KEY_MISSING'
      )
    }
    if (request.prompt.trim().length === 0) {
      throw new UniverError(
        'An image prompt is required: describe the image the user wants.',
        'INVALID_REQUEST'
      )
    }
    const gateway = await this.requireGateway()
    const worktreeId = await this.resolveEditBranch(gateway, request.file, request.worktreeId)
    await this.captureBaseline(gateway, request.file, worktreeId, request.unitId, signal)
    const source = await this.unitContent.renderSource(
      gateway,
      request.file,
      request.unitId,
      worktreeId,
      signal
    )
    if (source.unitType !== 'slide') {
      throw new UniverError('Only a Slide Unit can be image-edited.', 'UNIT_NOT_FOUND')
    }
    const element = locateSlideImage(source.unitData, {
      pageId: request.pageId,
      elementId: request.elementId
    })
    const reference = await readSlideImageBytes(element, {
      gateway,
      fileKey: fileKeyOf(request.file),
      worktreeId,
      timeoutMs: this.config.gatewayMutationTimeoutMs,
      ...(signal === undefined ? {} : { signal })
    })
    const generated = await generateImage({
      endpoint: this.config.imageEndpoint,
      apiKey,
      model: request.model ?? this.config.imageModel,
      prompt: request.prompt.trim(),
      image: reference,
      size: request.size ?? this.config.imageSize,
      timeoutMs: this.config.imageOperationTimeoutMs,
      ...(signal === undefined ? {} : { signal })
    })
    signal?.throwIfAborted()
    const execution = await this.unitContent.execute(
      gateway,
      request.file,
      replaceSlideImageCode(element, imageDataUrl(generated)),
      worktreeId,
      request.unitId,
      signal
    )
    this.invalidate(request.file, worktreeId)
    return {
      ok: true,
      operation: 'image-edit',
      file: request.file,
      result: {
        unitId: request.unitId,
        worktreeId,
        pageId: element.pageId,
        elementId: element.elementId,
        prompt: request.prompt.trim(),
        model: request.model ?? this.config.imageModel,
        reference: {
          mediaType: reference.mediaType,
          byteSize: reference.bytes.byteLength,
          width: element.width,
          height: element.height
        },
        generated: {
          mediaType: generated.mediaType,
          byteSize: generated.bytes.byteLength,
          width: generated.width,
          height: generated.height
        },
        execution: execution.result
      }
    }
  }

  /** Search or show the Facade API reference bundled for this plugin version. */
  apiReference(request: ApiReferenceRequest): Promise<UniverApiResult> {
    const result =
      request.action === 'find'
        ? this.api.find({
            terms: request.queries,
            ...(request.unit === undefined ? {} : { unit: request.unit }),
            ...(request.limit === undefined ? {} : { limit: request.limit })
          })
        : this.api.show(request.queries)
    return Promise.resolve({ ok: true, operation: 'api', result: result as unknown as JsonValue })
  }

  /** Search, read, export, or clear the bundled SVG resource library. */
  resources(
    request: ResourceOperationRequest,
    signal?: AbortSignal
  ): Promise<UniverResourceResult> {
    return this.resourceOperations.execute(request, signal)
  }

  /** Import one local deck into a draft worktree and hold it as the right panel's target. */
  async openPanel(request: OpenPanelSourceRequest, signal?: AbortSignal): Promise<PanelOpenResult> {
    await Promise.all([
      assertAuthorizedPath(request.workspace, request.file, false),
      assertAuthorizedPath(request.sourceWorkspace, request.source, true)
    ])
    const gateway = await this.requireGateway()
    const result = await this.panels.open(gateway, request, signal)
    if (result.ok) this.invalidate(request.file, result.panel.worktreeId)
    return result
  }

  /** Bind one already-authorized Unit to the right panel. */
  async bindPanel(request: BindPanelRequest, signal?: AbortSignal): Promise<PanelOpenResult> {
    await assertAuthorizedPath(request.workspace, request.file, true)
    const gateway = await this.requireGateway()
    return this.panels.bind(gateway, request, signal)
  }

  /** Release the session's panel binding. */
  releasePanel(sessionId: string): void {
    this.panels.release(sessionId)
  }

  /** The panel target bound to one session, if any. */
  panelOf(sessionId: string): OpenPanel | undefined {
    return this.panels.of(sessionId)
  }

  /**
   * Record the image the Viewer selected, so a later "edit this image" names the right element.
   *
   * The report carries no Unit: the panel already owns which Unit the browser shows, and taking it
   * from there keeps a stale report from attaching to a Unit the panel has moved away from.
   */
  reportPanelSelection(request: PanelSelectionRequest): void {
    const panel = this.panels.of(request.sessionId)
    if (panel === undefined) return
    if (request.pageId === undefined || request.elementId === undefined) {
      this.panels.select(request.sessionId, null)
      return
    }
    this.panels.select(request.sessionId, {
      unitId: panel.unitId,
      pageId: request.pageId,
      elementId: request.elementId
    })
  }

  /** Stop Gateway ownership and clear transient state. */
  async dispose(): Promise<void> {
    this.panels.clear()
    this.episodes.clear()
    this.stateCache.clear()
    this.unitCache.clear()
    try {
      // Runtime workers hold Gateway connections; close them before the Gateway stops.
      await this.unitContent.close()
    } finally {
      await this.gatewaySupervisor.dispose()
    }
  }

  /** Status value used by the Web Consumer. */
  async unitContentStatus(): Promise<'bundled'> {
    return 'bundled'
  }

  /**
   * Save one branch, writing a presentation back to the deck it was imported from.
   *
   * The branch is exported to a staged file beside the source first, so the presentation can still
   * be replaced only while the export succeeded, and the previous source is kept as `<source>.bak`
   * before the replacement. Merging last means a presentation that cannot be written — PowerPoint
   * holding it open, for example — leaves the unsaved branch intact instead of dropping it.
   */
  private async settlePresentation(
    gateway: string,
    target: {
      readonly workspace: WorkspacePath
      readonly file: UniverFilePath
      readonly worktreeId: WorktreeId
    }
  ): Promise<void> {
    try {
      const source = await this.panels.presentationOf(target.file)
      if (source === undefined) {
        await this.worktrees.action(gateway, target.file, target.worktreeId, 'merge')
        return
      }
      await assertAuthorizedPath(target.workspace, source, true)
      const api = new GatewayWorktreeApi(
        new GatewayClient(gateway, this.config.gatewayRequestTimeoutMs)
      )
      const slide = mapUnits(await api.listUnits(target.file, target.worktreeId)).find(
        (unit) => unit.type === SLIDE_UNIT_TYPE
      )
      if (slide === undefined) {
        await this.worktrees.action(gateway, target.file, target.worktreeId, 'merge')
        return
      }
      const staged = stagedPresentationPath(source)
      try {
        await this.unitContent.export(gateway, {
          workspace: target.workspace,
          file: target.file,
          unitId: brandedUnitId(slide.unitId),
          worktreeId: brandedWorktreeId(target.worktreeId),
          output: staged,
          outputWorkspace: target.workspace
        })
        await commitStagedPresentation(staged, source)
        await this.worktrees.action(gateway, target.file, target.worktreeId, 'merge')
        await markPresentationImported(target.file)
      } finally {
        await discardStagedPresentation(staged)
      }
    } finally {
      this.episodes.forgetWorktree(target.file, target.worktreeId)
    }
  }

  /**
   * Discard one branch, returning a presentation to the state the agent found it in.
   *
   * A plain discard would drop the user's own in-preview edits whenever the agent changed the deck
   * after them, because both live on the same branch. The snapshot taken before the agent's first
   * write is recreated in a fresh branch, which is why Discard reverts agent changes and keeps what
   * the user had already done.
   */
  private async discardPresentation(
    gateway: string,
    target: { readonly file: string; readonly worktreeId: string }
  ): Promise<void> {
    const snapshots = this.episodes.takeWorktree(target.file, target.worktreeId)
    await this.worktrees.action(gateway, target.file, target.worktreeId, 'discard')
    const restore = snapshots.filter(keepsUserEdits)
    if (restore.length === 0) return
    const api = new GatewayWorktreeApi(
      new GatewayClient(gateway, this.config.gatewayMutationTimeoutMs)
    )
    const created = await api.create(target.file, 'Restored before agent changes')
    requireGatewaySuccess(created, 'Gateway rejected the worktree creation.')
    const worktreeId = createdWorktreeId(created)
    for (const snapshot of restore) {
      const seeded = mapUnits(await api.listUnits(target.file, worktreeId)).find(
        (unit) => unit.unitId === snapshot.unitId
      )
      if (seeded !== undefined) {
        requireGatewaySuccess(
          await api.removeUnit(target.file, worktreeId, seeded.unitId),
          'Gateway rejected removing the saved presentation Unit.'
        )
      }
      requireGatewaySuccess(
        await api.createUnit(target.file, worktreeId, snapshot.kind, snapshot.name, snapshot.data),
        'Gateway rejected restoring the pre-agent presentation.'
      )
    }
  }

  /**
   * Remember one presentation Unit immediately before the agent's first write to it.
   *
   * Only a deck needs this: other content types have no user-facing "revert agent changes only"
   * contract, and snapshotting a large workbook on its first edit would cost a full export. The
   * branch revision is compared with the revision the branch was seeded from, so a deck the user
   * never touched is discarded normally instead of being recreated as an identical Unit.
   */
  private async captureBaseline(
    gateway: string,
    file: string,
    worktreeId: string,
    unitId: string,
    signal?: AbortSignal
  ): Promise<void> {
    if (this.episodes.has(file, worktreeId, unitId)) return
    const source = await this.panels.presentationOf(file)
    if (source === undefined) return
    const client = new GatewayClient(gateway, this.config.gatewayRequestTimeoutMs)
    const [units, worktrees] = await Promise.all([
      new GatewayWorktreeApi(client).listUnits(file, worktreeId),
      new GatewayFileApi(client).listWorktrees(file)
    ])
    const unit = mapUnits(units).find((candidate) => candidate.unitId === unitId)
    if (unit === undefined) return
    const branch = mapWorktrees(worktrees).find((entry) => entry.worktreeId === worktreeId)
    // A Unit the branch created itself starts at revision 1; a Unit seeded from trunk starts at the
    // trunk revision the branch recorded.
    const seededRev = branch?.baseline[unitId] ?? 1
    if (unit.headRev === seededRev) {
      // Nothing was edited in the preview before the agent: a plain branch discard restores it.
      this.episodes.mark(file, worktreeId, unitId)
      return
    }
    const captured = await this.unitContent.renderSource(gateway, file, unitId, worktreeId, signal)
    this.episodes.record(file, worktreeId, {
      unitId,
      name: unit.name,
      kind: presentationKind(unit.type),
      data: captured.unitData,
      seededRev,
      capturedRev: unit.headRev
    })
  }

  private async computeFileState(file: string): Promise<FileState> {
    let status = await this.gatewaySupervisor.status()
    if (status.gateway === null && this.config.autoStartGateway) {
      const started = await this.gatewaySupervisor.ensure()
      if (started.ok)
        status = { phase: 'running', gateway: started.gateway, owned: !started.reused }
    }
    if (status.gateway === null)
      throw new UniverError(
        status.reason ?? 'Univer Gateway is not available.',
        'GATEWAY_UNAVAILABLE'
      )
    const gateway = status.gateway
    const records = mapWorktrees(
      await new GatewayFileApi(
        new GatewayClient(gateway, this.config.gatewayRequestTimeoutMs)
      ).listWorktrees(file)
    )
    return {
      ok: true,
      file,
      gateway,
      gatewayRunning: true,
      viewerUrl: `${VIEWER_BASE}/?file=${encodeURIComponent(fileKeyOf(file))}`,
      branch: await this.branchState(gateway, file, newestActive(records))
    }
  }

  /**
   * Project the file's one editable branch.
   *
   * Only the newest active worktree is projected: older ones are leftovers the browser must not
   * surface as extra branches. Everything the preview needs is computed here — the Unit list it can
   * switch between, each Unit's Viewer target, and whether the branch still differs from the last
   * saved version, which is what enables Save and Discard.
   */
  private async branchState(
    gateway: string,
    file: string,
    record: ReturnType<typeof mapWorktrees>[number] | undefined
  ): Promise<FileBranch | null> {
    if (record === undefined) return null
    const base = `${VIEWER_BASE}/?file=${encodeURIComponent(fileKeyOf(file))}`
    const worktreeUrl = `${base}&worktree=${encodeURIComponent(record.worktreeId)}&mode=embedded&scope=worktree`
    const client = new GatewayClient(gateway, this.config.gatewayRequestTimeoutMs)
    const units = mapUnits(await new GatewayWorktreeApi(client).listUnits(file, record.worktreeId))
    const changed = await this.unitCache.get(`${file}\u0000${record.worktreeId}`, () =>
      this.worktrees.changedUnits(gateway, file, record.worktreeId)
    )
    const kinds = new Map(changed.map((unit) => [unit.unitId, unit.kind]))
    return {
      worktreeId: record.worktreeId,
      name: record.name,
      unsaved: changed.length > 0,
      worktreeUrl,
      units: units.map((unit) => ({
        unitId: unit.unitId,
        name: unit.name,
        type: unitKind(unit.type),
        kind: kinds.get(unit.unitId) ?? 'unchanged',
        worktreeUrl: `${worktreeUrl}&unit=${encodeURIComponent(unit.unitId)}`
      }))
    }
  }

  /**
   * Resolve the branch an edit request writes into.
   *
   * An explicit `worktreeId` still wins, so a caller that already knows the branch keeps its exact
   * target; a legacy `ready` branch is reopened because this model has no separate submission step,
   * and edits belong in a writable branch. Everything else edits the file's single branch.
   */
  private async resolveEditBranch(
    gateway: string,
    file: string,
    explicit: WorktreeId | undefined
  ): Promise<string> {
    if (explicit === undefined) return this.ensureEditBranch(gateway, file)
    const client = new GatewayClient(gateway, this.config.gatewayMutationTimeoutMs)
    const record = mapWorktrees(await new GatewayFileApi(client).listWorktrees(file)).find(
      (entry) => entry.worktreeId === explicit
    )
    if (record === undefined)
      throw new UniverError(`Worktree ${explicit} was not found.`, 'WORKTREE_NOT_FOUND')
    if (record.status === 'draft') return record.worktreeId
    if (record.status !== 'ready')
      throw new UniverError(
        `Worktree ${explicit} is ${record.status}; it no longer accepts edits.`,
        'WORKTREE_STATUS_INVALID'
      )
    requireGatewaySuccess(
      await new GatewayWorktreeApi(client).action(file, explicit, 'reopen'),
      'Gateway rejected reopening the edit branch.'
    )
    this.invalidate(file, explicit)
    return explicit
  }

  /** The file's one editable branch: the newest active worktree, or a fresh draft beside the saved version. */
  private async ensureEditBranch(gateway: string, file: string, name?: string): Promise<string> {
    const client = new GatewayClient(gateway, this.config.gatewayMutationTimeoutMs)
    const existing = newestActive(
      mapWorktrees(await new GatewayFileApi(client).listWorktrees(file))
    )
    if (existing !== undefined)
      return this.resolveEditBranch(gateway, file, brandedWorktreeId(existing.worktreeId))
    const api = new GatewayWorktreeApi(client)
    const created = await api.create(file, name ?? 'Univer edits')
    requireGatewaySuccess(created, 'Gateway rejected the worktree creation.')
    const id = brandedWorktreeId(createdWorktreeId(created))
    this.invalidate(file, id)
    return id
  }

  private invalidate(file: string, worktreeId: string): void {
    this.stateCache.delete(file)
    this.unitCache.delete(`${file}\u0000${worktreeId}`)
  }

  private async requireGateway(): Promise<string> {
    const available = await this.ensureGateway()
    if (!available.ok) throw new UniverError(available.reason, 'GATEWAY_UNAVAILABLE')
    return available.gateway
  }
}

/** Public kind of one Gateway Unit type, for an episode snapshot. */
function presentationKind(type: number): UniverUnitKind {
  if (type === 1) return 'doc'
  if (type === 2) return 'sheet'
  if (type === 3) return 'slide'
  if (type === 5) return 'base'
  return 'board'
}

function unitResult(unit: ReturnType<typeof mapUnits>[number]): { [key: string]: JsonValue } {
  return {
    unitId: unit.unitId,
    name: unit.name,
    kind: unitKind(unit.type),
    type: unit.type,
    headRevision: unit.headRev
  }
}

function worktreeResult(worktree: ReturnType<typeof mapWorktrees>[number]): {
  [key: string]: JsonValue
} {
  return {
    worktreeId: worktree.worktreeId,
    name: worktree.name,
    status: worktree.status,
    baseline: worktree.baseline,
    ...(worktree.createdAt === undefined ? {} : { createdAt: worktree.createdAt }),
    ...(worktree.mergedAt === undefined ? {} : { mergedAt: worktree.mergedAt })
  }
}

/** The only branch the browser shows: newest of the worktrees that still accept work. */
function newestActive(
  worktrees: readonly ReturnType<typeof mapWorktrees>[number][]
): ReturnType<typeof mapWorktrees>[number] | undefined {
  return worktrees
    .filter((entry) => entry.status === 'draft' || entry.status === 'ready')
    .sort((left, right) => (right.createdAt ?? '').localeCompare(left.createdAt ?? ''))[0]
}

function createdWorktreeId(value: JsonValue): string {
  const record = asRecord(value)
  const worktreeId = record.worktreeId
  if (typeof worktreeId === 'string' && worktreeId.length > 0) return worktreeId
  throw new UniverError(
    'Gateway worktree result is missing worktreeId.',
    'GATEWAY_INVALID_RESPONSE'
  )
}

function requireGatewaySuccess(value: JsonValue, fallback: string): void {
  if (isRecord(value) && isRecord(value.error) && value.error.code === 0) {
    throw new UniverError(gatewayErrorMessage(value) ?? fallback, 'GATEWAY_REQUEST_REJECTED')
  }
  if (isRecord(value) && value.ok === false) {
    throw new UniverError(gatewayErrorMessage(value) ?? fallback, 'GATEWAY_REQUEST_REJECTED')
  }
}

function asRecord(value: JsonValue): { [key: string]: JsonValue } {
  if (!isRecord(value))
    throw new UniverError('Gateway result must be an object.', 'GATEWAY_INVALID_RESPONSE')
  return value
}
