import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { unitId, worktreeId } from '../../service/identifiers.ts'
import { operationOutput, operationTitle } from '../presentation.ts'
import { assertPptxExportAllowed } from '../open-presentation.ts'
import { existingToolFile, newToolPath } from '../workspace.ts'

/** Create the `univer_export` tool definition. */
export function exportTool(ctx: Context, timeoutMs: number) {
  return defineTool({
    name: 'univer_export',
    description:
      'Export a .univer document or unit to a user-facing file format. While a presentation is open, do not export another .pptx: the preview Save writes that same file.',
    timeoutMs,
    parameters: {
      file: {
        type: 'string',
        required: true,
        description: 'Workspace-relative or absolute .univer path.'
      },
      output: {
        type: 'string',
        required: true,
        description: 'Workspace-relative or absolute output file path.'
      },
      unitId: {
        type: 'string',
        required: true,
        description: 'Explicit Unit id from univer_status.'
      },
      worktreeId: { type: 'string', description: 'Optional worktree scope; omit to export trunk.' }
    },
    output: operationOutput,
    async execute(args, exec) {
      const [target, output] = await Promise.all([
        existingToolFile(exec, args.file),
        newToolPath(exec, args.output)
      ])
      // Save in the preview writes the original pptx. A second pptx is a new file the preview opens.
      assertPptxExportAllowed(ctx, exec, output.path)
      return ctx.univer.exportUnitContent(
        {
          workspace: target.workspace,
          file: target.path,
          outputWorkspace: output.workspace,
          output: output.path,
          unitId: unitId(args.unitId),
          ...(args.worktreeId === undefined ? {} : { worktreeId: worktreeId(args.worktreeId) })
        },
        exec.signal
      )
    },
    presentCall: (args) => ({
      card: 'generic',
      title: operationTitle('export', args.file),
      kind: 'execute'
    })
  })
}
