import type { WorktreeReviewAction, WorktreeActionResult } from '../../shared/wire/actions.ts'
import type { OpenPanel, PanelOpenResult } from '../../shared/wire/panel.ts'
import type { FileState } from '../../shared/wire/state.ts'
import type { EnsureGatewayResult, GatewayStatus } from '../../shared/wire/status.ts'
import type { UniverFilePath, UnitId, WorkspacePath, WorktreeId } from './identifiers.ts'

/** JSON values accepted across the model tool boundary. */
export type JsonValue =
  | null
  | boolean
  | number
  | string
  | JsonValue[]
  | { [key: string]: JsonValue }

/** Unit kinds supported by the collaboration Gateway. */
export type UniverUnitKind = 'sheet' | 'doc' | 'slide' | 'base' | 'board'

/** File request tied to one authorized workspace. */
export interface ScopedFileRequest {
  readonly workspace: WorkspacePath
  readonly file: UniverFilePath
}

/** Request for one file's collaboration state. */
export interface FileStateRequest extends ScopedFileRequest {}

/** Request for a model-readable file status projection. */
export interface FileStatusRequest extends ScopedFileRequest {
  readonly worktreeId?: WorktreeId
  readonly unitId?: UnitId
}

/** Request for creating an empty Univer file. */
export interface NewUniverFileRequest extends ScopedFileRequest {}

/** Request for a browser worktree lifecycle decision. */
export interface WorktreeActionRequest extends ScopedFileRequest {
  readonly action: WorktreeReviewAction
  readonly worktreeId: WorktreeId
}

/**
 * Agent-facing worktree request.
 *
 * `create` no longer means "another branch": it ensures the file's one editable branch and reports
 * its identity. Save and Discard name the branch the model already knows from `univer_status`.
 */
export type WorktreeOperationRequest = ScopedFileRequest &
  (
    | { readonly action: 'create'; readonly name?: string }
    | { readonly action: 'merge' | 'discard'; readonly worktreeId: WorktreeId }
  )

/**
 * The branch every edit request resolves to.
 *
 * Omitting `worktreeId` is the normal case: the Host edits the file's single branch, which is the
 * one the browser shows, so the model never has to pick between branches.
 */
export type EditWorktreeTarget = { readonly worktreeId?: WorktreeId }

/** Draft-worktree Unit lifecycle request. */
export type UnitOperationRequest = ScopedFileRequest &
  EditWorktreeTarget &
  (
    | {
        readonly action: 'create'
        readonly kind: UniverUnitKind
        readonly name: string
      }
    | {
        readonly action: 'remove'
        readonly unitId: UnitId
      }
  )

/** Request for inspecting one Unit. */
export interface InspectUnitContentRequest extends ScopedFileRequest {
  readonly elementIds?: readonly string[]
  readonly unitId: UnitId
  readonly range?: string
  readonly worktreeId?: WorktreeId
}

/** Request for executing Univer Facade code. */
export interface ExecuteUnitContentRequest extends ScopedFileRequest, EditWorktreeTarget {
  readonly code: string
  readonly unitId: UnitId
}

/** Output size the image model is asked for. */
export type ImageOutputSize = '1K' | '2K' | '4K'

/**
 * Request for replacing one selected Slide image with a model-generated one.
 *
 * The browser reports which image the user picked, so the tool resolves `pageId` and `elementId`
 * from the panel selection and names them explicitly here: the Provider never reads browser state.
 */
export interface EditSlideImageRequest extends ScopedFileRequest, EditWorktreeTarget {
  readonly unitId: UnitId
  readonly pageId: string
  readonly elementId: string
  /** The user's own wording; it is the model prompt, so it must describe the wanted image. */
  readonly prompt: string
  readonly model?: string
  readonly size?: ImageOutputSize
}

/** Compatibility mode for .doc/.docx imports; traditional keeps the source page setup. */
export type ImportDocType = 'traditional' | 'modern'

/** Request for importing one Office file as a Unit in the editable branch. */
export interface ImportUnitContentRequest extends ScopedFileRequest, EditWorktreeTarget {
  readonly source: string
  readonly sourceWorkspace: WorkspacePath
  readonly name: string
  readonly docType?: ImportDocType
}

/** Request for exporting one Unit. */
export interface ExportUnitContentRequest extends ScopedFileRequest {
  readonly output: string
  readonly outputWorkspace: WorkspacePath
  readonly unitId: UnitId
  readonly worktreeId?: WorktreeId
}

/** Request for importing a local Office file into the right panel's `.univer` container. */
export interface OpenPanelSourceRequest extends ScopedFileRequest {
  /** Authorized existing source, such as a `.pptx` inside the session workspace. */
  readonly source: string
  readonly sourceWorkspace: WorkspacePath
  /** DSH session whose right panel opened the file. */
  readonly sessionId: string
}

/** Request for binding an already-authorized Unit to the right panel. */
export interface BindPanelRequest extends ScopedFileRequest {
  readonly sessionId: string
  /** Omit to bind the newest draft worktree holding a viewable Unit. */
  readonly worktreeId?: WorktreeId
  /** Omit to bind a presentation Unit, or else the first Unit in the selected worktree. */
  readonly unitId?: UnitId
  readonly source?: string
}

/** Request for deterministic Slide layout analysis. */
export interface LintUnitLayoutRequest extends ScopedFileRequest {
  readonly unitId: UnitId
  readonly worktreeId?: WorktreeId
  readonly pages?: readonly (number | string)[]
}

/** Unit-specific viewport requested from the browser screenshot runtime. */
export type ScreenshotTarget =
  | {
      readonly kind: 'sheet-range'
      readonly range: string
      readonly sheetName?: string
      readonly scale?: number
    }
  | {
      readonly kind: 'paged-unit'
      readonly pages?: readonly (number | string)[]
      readonly contactSheet?: {
        readonly tile?: { readonly columns: number; readonly rows: number }
      }
      readonly scale?: number
    }
  | {
      readonly kind: 'board-content'
      readonly elementIds?: readonly string[]
      readonly padding?: number
      readonly region?: {
        readonly left: number
        readonly top: number
        readonly width: number
        readonly height: number
      }
      readonly scale?: number
    }
  | {
      readonly kind: 'unit-viewport'
      readonly scale: number
    }

/** Request for rendering one Unit into workspace PNG files. */
export interface ScreenshotUnitRequest extends ScopedFileRequest {
  readonly unitId: UnitId
  readonly worktreeId?: WorktreeId
  readonly output: string
  readonly outputWorkspace: WorkspacePath
  readonly target?: ScreenshotTarget
}

/** Request for printing one Unit to an authorized workspace PDF file. */
export interface PrintPdfUnitRequest extends ScopedFileRequest {
  readonly unitId: UnitId
  readonly worktreeId?: WorktreeId
  readonly output: string
  readonly outputWorkspace: WorkspacePath
}

/** One PNG returned to the Tools Consumer before attachment persistence. */
export interface ScreenshotServiceImage {
  readonly data: string
  readonly height: number
  readonly mediaType: 'image/png'
  readonly name: string
  readonly path: string
  readonly width: number
  readonly metadata: JsonValue
}

/** Browser screenshot result with bytes encoded for the same-process Service boundary. */
export interface ScreenshotServiceResult {
  readonly ok: true
  readonly operation: 'screenshot'
  readonly file: string
  readonly result: {
    readonly unitId: string
    readonly unitType: UniverUnitKind
    readonly images: readonly ScreenshotServiceImage[]
  }
}

/** Resource-library operation over the bundled multi-registry manifest. */
export type ResourceOperationRequest =
  | { readonly action: 'registries' }
  | {
      readonly action: 'find'
      readonly queries: readonly string[]
      readonly registries?: readonly string[]
      readonly limit?: number
    }
  | { readonly action: 'read'; readonly handle: string }
  | {
      readonly action: 'export'
      readonly handles: readonly string[]
      readonly output: string
      readonly outputWorkspace: WorkspacePath
    }
  | { readonly action: 'clear-cache' }

/** Structured resource result that does not expose Provider cache paths. */
export interface UniverResourceResult {
  readonly ok: true
  readonly operation: 'resources'
  readonly result: JsonValue
}

/** Request for compiling one SVG and applying it to a Slide page. */
export interface CompileSvgRequest extends ScopedFileRequest, EditWorktreeTarget {
  readonly source: string
  readonly sourceWorkspace: WorkspacePath
  readonly unitId: UnitId
  readonly page: number
  readonly mode?: 'replace' | 'add'
}

/** Version-matched Facade reference lookup. */
export type ApiReferenceRequest =
  | {
      readonly action: 'find'
      readonly queries: readonly string[]
      readonly unit?: UniverUnitKind
      readonly limit?: number
    }
  | { readonly action: 'show'; readonly queries: readonly string[] }

/** Structured operation result logged in the DSH session. */
export interface UniverOperationResult {
  readonly ok: true
  readonly operation:
    | 'new'
    | 'status'
    | 'inspect'
    | 'execute'
    | 'import'
    | 'export'
    | 'lint'
    | 'screenshot'
    | 'print-pdf'
    | 'compile-svg'
    | 'image-edit'
    | 'unit'
    | 'worktree'
  readonly file: string
  readonly result: JsonValue
}

/** Structured API-reference result logged in the DSH session. */
export interface UniverApiResult {
  readonly ok: true
  readonly operation: 'api'
  readonly result: JsonValue
}

/** Stable methods offered by the Univer service. */
export interface UniverServiceMethods {
  gatewayStatus(): Promise<GatewayStatus>
  ensureGateway(): Promise<EnsureGatewayResult>
  unitContentStatus(): Promise<'bundled' | 'unavailable'>
  fileState(request: FileStateRequest): Promise<FileState>
  worktreeAction(request: WorktreeActionRequest): Promise<WorktreeActionResult>
  newFile(request: NewUniverFileRequest, signal?: AbortSignal): Promise<UniverOperationResult>
  status(request: FileStatusRequest, signal?: AbortSignal): Promise<UniverOperationResult>
  worktree(request: WorktreeOperationRequest, signal?: AbortSignal): Promise<UniverOperationResult>
  unit(request: UnitOperationRequest, signal?: AbortSignal): Promise<UniverOperationResult>
  inspectUnitContent(
    request: InspectUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult>
  executeUnitContent(
    request: ExecuteUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult>
  importUnitContent(
    request: ImportUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult>
  exportUnitContent(
    request: ExportUnitContentRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult>
  lintUnitLayout(
    request: LintUnitLayoutRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult>
  screenshotUnit(
    request: ScreenshotUnitRequest,
    signal?: AbortSignal
  ): Promise<ScreenshotServiceResult>
  printUnitPdf(request: PrintPdfUnitRequest, signal?: AbortSignal): Promise<UniverOperationResult>
  compileSvg(request: CompileSvgRequest, signal?: AbortSignal): Promise<UniverOperationResult>
  /** Replace one selected Slide image with a model-generated image; needs an Ark API key. */
  editSlideImage(
    request: EditSlideImageRequest,
    signal?: AbortSignal
  ): Promise<UniverOperationResult>
  apiReference(request: ApiReferenceRequest): Promise<UniverApiResult>
  resources(request: ResourceOperationRequest, signal?: AbortSignal): Promise<UniverResourceResult>
  /** Import a local source into a draft worktree and hold it as the right panel's target. */
  openPanel(request: OpenPanelSourceRequest, signal?: AbortSignal): Promise<PanelOpenResult>
  /** Hold an already-authorized Unit as the right panel's target. */
  bindPanel(request: BindPanelRequest, signal?: AbortSignal): Promise<PanelOpenResult>
  /** Release the session's panel binding; safe when nothing is bound. */
  releasePanel(sessionId: string): void
  /** The panel target currently bound to one session, if any. */
  panelOf(sessionId: string): OpenPanel | undefined
  /** Record the image the browser has selected inside the session's bound Unit. */
  reportPanelSelection(request: PanelSelectionRequest): void
}

/** Body of one selection report from the Viewer. */
export interface PanelSelectionRequest {
  readonly sessionId: string
  readonly pageId?: string
  readonly elementId?: string
}
