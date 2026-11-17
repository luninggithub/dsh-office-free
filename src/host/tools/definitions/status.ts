import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import type { OpenPanel } from '../../../shared/wire/panel.ts'
import { isRecord } from '../../adapters/gateway/mapping.ts'
import { unitId, worktreeId } from '../../service/identifiers.ts'
import type { JsonValue, UniverOperationResult } from '../../service/types.ts'
import { operationOutput, operationTitle } from '../presentation.ts'
import { existingToolFile } from '../workspace.ts'

/** Create the `univer_status` tool definition. */
export function statusTool(ctx: Context, timeoutMs: number) {
  return defineTool({
    name: 'univer_status',
    description:
      'List a .univer file’s saved Units and its single edit branch. Call this before choosing unitId or continuing prior work: edits always go to the branch reported here, and its worktreeId is only needed by univer_worktree merge/discard. When the DSH client holds a target in the right Univer panel, the result also reports it as openPanel: that file, worktreeId, and unitId are the active edit target.',
    timeoutMs,
    parameters: {
      file: {
        type: 'string',
        required: true,
        description: 'Workspace-relative or absolute .univer path.'
      },
      worktreeId: {
        type: 'string',
        description: 'Optional worktree whose Units should be returned.'
      },
      unitId: { type: 'string', description: 'Optional Unit filter.' }
    },
    output: operationOutput,
    async execute(args, exec) {
      const target = await existingToolFile(exec, args.file)
      const result = await ctx.univer.status(
        {
          workspace: target.workspace,
          file: target.path,
          ...(args.worktreeId === undefined ? {} : { worktreeId: worktreeId(args.worktreeId) }),
          ...(args.unitId === undefined ? {} : { unitId: unitId(args.unitId) })
        },
        exec.signal
      )
      return withOpenPanel(ctx, result, exec.agent?.session.header.id)
    },
    presentCall: (args) => ({
      card: 'generic',
      title: operationTitle('status', args.file),
      kind: 'read'
    })
  })
}

/**
 * Add the right panel's current target to the status result.
 *
 * The panel binding is browser state, so the model cannot infer it from the file alone; the
 * structured status result is its only replayable channel.
 */
function withOpenPanel(
  ctx: Context,
  result: UniverOperationResult,
  sessionId: string | undefined
): UniverOperationResult {
  const panel = sessionId === undefined ? undefined : ctx.univer.panelOf(sessionId)
  if (panel === undefined) return result
  const detail = isRecord(result.result) ? result.result : {}
  return {
    ...result,
    result: { ...detail, openPanel: panelRecord(panel) }
  }
}

/**
 * Project one panel binding for the model.
 *
 * `selection` is the image the user last picked in the Viewer; it is what `univer_image_edit`
 * defaults to, so the model can confirm the target before rewriting a picture.
 */
function panelRecord(panel: OpenPanel): { [key: string]: JsonValue } {
  return {
    source: panel.source,
    file: panel.file,
    worktreeId: panel.worktreeId,
    unitId: panel.unitId,
    selection:
      panel.selection === null
        ? null
        : {
            unitId: panel.selection.unitId,
            pageId: panel.selection.pageId,
            elementId: panel.selection.elementId
          }
  }
}
