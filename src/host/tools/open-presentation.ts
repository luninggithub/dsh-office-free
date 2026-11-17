import type { Context } from '@deepseek-ai/cordis'
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'
import type { OpenPanel } from '../../shared/wire/panel.ts'
import { UniverError } from '../service/errors.ts'
import type { UniverService } from '../service/univer-service.ts'

/**
 * The deck a session already has open, when the user opened a `.pptx`.
 *
 * A `.univer` opened directly has no source, so it is not locked: only a presentation
 * the user clicked has an original file that later edits must keep.
 */
export function openPresentationOf(ctx: Context, exec: ToolRunContext): OpenPanel | undefined {
  const sessionId = exec.agent?.session.header.id
  if (sessionId === undefined) return undefined
  const panel = readPanel(ctx, sessionId)
  if (panel?.source == null || panel.source.length === 0) return undefined
  return panel
}

/**
 * Refuse a write that targets some file other than the open deck.
 *
 * Image generation, SVG compilation, and new Slide Units belong on the deck the
 * user is looking at. Pointing them at another container is how a second presentation
 * gets created and then opened beside the original.
 */
export function assertSameOpenPresentation(ctx: Context, exec: ToolRunContext, file: string): void {
  const panel = openPresentationOf(ctx, exec)
  if (panel === undefined) return
  if (samePath(panel.file, file) || samePath(panel.source ?? '', file)) return
  throw locked(panel)
}

/** Refuse to import another `.pptx` over, or beside, the deck already open. */
export function assertPptxImportAllowed(ctx: Context, exec: ToolRunContext, source: string): void {
  const panel = openPresentationOf(ctx, exec)
  if (panel === undefined || !isPresentationPath(source)) return
  throw locked(panel)
}

/**
 * Refuse to export any `.pptx` while a deck is open.
 *
 * The preview's Save writes the open deck back to its original file. Exporting a
 * `.pptx`, including onto that same path, writes a presentation outside that save.
 */
export function assertPptxExportAllowed(ctx: Context, exec: ToolRunContext, output: string): void {
  const panel = openPresentationOf(ctx, exec)
  if (panel === undefined || !isPresentationPath(output)) return
  throw locked(panel)
}

function isPresentationPath(path: string): boolean {
  return /\.pptx$/i.test(path)
}

/** `UniverService` is referenced so its module, which augments `Context`, stays in this file. */
function readPanel(ctx: Context, sessionId: string): ReturnType<UniverService['panelOf']> {
  return ctx.univer.panelOf(sessionId)
}

function locked(panel: OpenPanel): UniverError {
  return new UniverError(
    `A presentation is already open at ${panel.source}. Edit that deck in place with univer_execute, univer_compile_svg, or univer_image_edit. Saving in the preview writes the same file. Do not create, import, or export another presentation.`,
    'PRESENTATION_ALREADY_OPEN'
  )
}

/** Same identity the Client uses for preview keys: `/` separators, case-folded on Windows. */
function samePath(left: string, right: string): boolean {
  return identity(left) === identity(right)
}

function identity(path: string): string {
  const slashed = path.replace(/\\/g, '/')
  const windows = slashed.startsWith('//') || /^[A-Za-z]:\//.test(slashed)
  return windows ? slashed.toLowerCase() : slashed
}
