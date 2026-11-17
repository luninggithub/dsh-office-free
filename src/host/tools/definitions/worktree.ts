import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { worktreeId } from '../../service/identifiers.ts'
import { UniverError } from '../../service/errors.ts'
import { operationOutput, operationTitle } from '../presentation.ts'
import { stripGatewaySuccessEnvelope } from '../normalize.ts'
import { existingToolFile } from '../workspace.ts'

/**
 * Create the agent-safe `univer_worktree` tool definition.
 *
 * There is one branch per file. `create` ensures it and reports its id; `merge` is the user's Save
 * and `discard` is theirs to reject, so both stay behind approval. Editing needs neither: every
 * `univer_execute` / `univer_unit` / `univer_import` call resolves to the same branch.
 */
export function worktreeTool(ctx: Context, timeoutMs: number) {
  return defineTool({
    name: 'univer_worktree',
    description:
      'Manage the file’s single edit branch. `create` ensures that branch exists and returns its worktreeId; `merge` saves it into the current version, writing a presentation back to the .pptx it was opened from and keeping the previous version as a .bak beside it; `discard` throws it away. Merge and discard require user approval.',
    timeoutMs,
    parameters: {
      action: {
        type: 'string',
        required: true,
        enum: ['create', 'merge', 'discard'],
        description: 'Branch action.'
      },
      file: {
        type: 'string',
        required: true,
        description: 'Workspace-relative or absolute .univer path.'
      },
      worktreeId: {
        type: 'string',
        description: 'Optional branch id from univer_status. Omit to use the file’s only branch.'
      },
      name: { type: 'string', description: 'Optional human-readable name used by create.' }
    },
    output: operationOutput,
    async execute(args, exec) {
      const target = await existingToolFile(exec, args.file)
      if (args.action === 'create') {
        return stripGatewaySuccessEnvelope(
          await ctx.univer.worktree(
            {
              action: 'create',
              workspace: target.workspace,
              file: target.path,
              ...(args.name === undefined ? {} : { name: args.name })
            },
            exec.signal
          )
        )
      }
      if (args.worktreeId === undefined || args.worktreeId.length === 0) {
        throw new UniverError(
          `univer_worktree ${args.action} requires the worktreeId reported by univer_status.`,
          'INVALID_REQUEST'
        )
      }
      return stripGatewaySuccessEnvelope(
        await ctx.univer.worktree(
          {
            action: args.action,
            workspace: target.workspace,
            file: target.path,
            worktreeId: worktreeId(args.worktreeId)
          },
          exec.signal
        )
      )
    },
    presentCall: (args) => ({
      card: 'generic',
      title: operationTitle(`worktree ${args.action}`, args.file),
      kind: 'execute'
    })
  })
}
