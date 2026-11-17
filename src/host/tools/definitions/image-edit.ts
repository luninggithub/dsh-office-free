import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import type { OpenPanel, PanelImageSelection } from '../../../shared/wire/panel.ts'
import { UniverError } from '../../service/errors.ts'
import { unitId, worktreeId } from '../../service/identifiers.ts'
import { operationOutput, operationTitle } from '../presentation.ts'
import { existingToolFile } from '../workspace.ts'

/** Target the edit runs against, after the panel fills in whatever the caller omitted. */
interface ImageEditTarget {
  readonly unitId: string
  readonly pageId: string
  readonly elementId: string
}

/** Create the `univer_image_edit` tool definition. */
export function imageEditTool(ctx: Context, timeoutMs: number) {
  return defineTool({
    name: 'univer_image_edit',
    description:
      'Replace one image on a Slide with a newly generated one. When the user selects an image in the right panel and asks to replace or edit “this image”, omit file, unitId, pageId, and elementId: the panel supplies them. The prompt is passed to the image model verbatim, so restate the user’s request as an image instruction.',
    timeoutMs,
    parameters: {
      prompt: {
        type: 'string',
        required: true,
        description:
          'What the generated image should look like, in the user’s words, e.g. “改成水彩风格”.'
      },
      file: {
        type: 'string',
        description: 'Optional .univer path; defaults to the file open in the right panel.'
      },
      unitId: {
        type: 'string',
        description: 'Optional Slide Unit id; defaults to the panel Unit.'
      },
      pageId: {
        type: 'string',
        description: 'Optional slide page id; defaults to the image selected in the panel.'
      },
      elementId: {
        type: 'string',
        description: 'Optional image element id; defaults to the image selected in the panel.'
      },
      worktreeId: {
        type: 'string',
        description: 'Optional branch id; omit to write the file’s single edit branch.'
      },
      model: {
        type: 'string',
        description: 'Optional image model id; defaults to the Host config.'
      },
      size: {
        type: 'string',
        enum: ['1K', '2K', '4K'],
        description: 'Optional output size class; defaults to the Host config.'
      }
    },
    output: operationOutput,
    async execute(args, exec) {
      const panel = panelOf(ctx, exec.agent?.session.header.id)
      const selection = requireSelection(args, panel)
      const file = args.file ?? panel?.file
      if (file === undefined) {
        throw new UniverError(
          'No image target: open a presentation in the right panel or pass file, unitId, pageId, and elementId.',
          'IMAGE_SELECTION_MISSING'
        )
      }
      const target = await existingToolFile(exec, file)
      const resolved = resolveTarget(args, panel, selection, target.path)
      return ctx.univer.editSlideImage(
        {
          workspace: target.workspace,
          file: target.path,
          unitId: unitId(resolved.unitId),
          pageId: resolved.pageId,
          elementId: resolved.elementId,
          prompt: args.prompt,
          ...(args.worktreeId === undefined ? {} : { worktreeId: worktreeId(args.worktreeId) }),
          ...(args.model === undefined ? {} : { model: args.model }),
          ...(args.size === undefined ? {} : { size: args.size })
        },
        exec.signal
      )
    },
    presentCall: (args) => ({
      card: 'generic',
      title: operationTitle('image edit', args.file ?? 'selected image'),
      kind: 'execute'
    })
  })
}

function panelOf(ctx: Context, sessionId: string | undefined): OpenPanel | undefined {
  return sessionId === undefined ? undefined : ctx.univer.panelOf(sessionId)
}

/**
 * Find the image this edit applies to.
 *
 * Explicit ids win over the panel, but the panel's selection is the only source when the user just
 * points at a picture, so a caller that names neither fails with the actionable reason instead of
 * editing an arbitrary image.
 */
function requireSelection(
  args: {
    readonly pageId?: string
    readonly elementId?: string
  },
  panel: OpenPanel | undefined
): PanelImageSelection | undefined {
  const hasPage = args.pageId !== undefined
  const hasElement = args.elementId !== undefined
  if (hasPage !== hasElement) {
    throw new UniverError(
      'pageId and elementId must be provided together.',
      'IMAGE_SELECTION_INCOMPLETE'
    )
  }
  if (hasPage) return undefined
  return panel?.selection ?? undefined
}

function resolveTarget(
  args: {
    readonly unitId?: string
    readonly pageId?: string
    readonly elementId?: string
  },
  panel: OpenPanel | undefined,
  selection: PanelImageSelection | undefined,
  file: string
): ImageEditTarget {
  const unit = args.unitId ?? panel?.unitId
  if (unit === undefined) {
    throw new UniverError(
      `No Slide Unit for ${file}: pass unitId from univer_status.`,
      'UNIT_NOT_FOUND'
    )
  }
  if (args.pageId !== undefined && args.elementId !== undefined) {
    return { unitId: unit, pageId: args.pageId, elementId: args.elementId }
  }
  if (selection === undefined) {
    throw new UniverError(
      'No image is selected: select an image on a slide, then ask again.',
      'IMAGE_SELECTION_MISSING'
    )
  }
  // A selection names its own Unit, so a report left over from another target cannot be applied to
  // the Unit this call writes.
  if (selection.unitId !== unit) {
    throw new UniverError(
      'The selected image belongs to another Unit; select an image in the Unit you are editing.',
      'IMAGE_SELECTION_STALE'
    )
  }
  return { unitId: unit, pageId: selection.pageId, elementId: selection.elementId }
}
