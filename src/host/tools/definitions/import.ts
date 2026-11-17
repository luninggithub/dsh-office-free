import { defineTool } from '@deepseek-ai/dsh-tools'
import type { Context } from '@deepseek-ai/cordis'
import { UniverError } from '../../service/errors.ts'
import { worktreeId } from '../../service/identifiers.ts'
import { operationOutput, operationTitle } from '../presentation.ts'
import { stripGatewaySuccessEnvelope } from '../normalize.ts'
import { assertPptxImportAllowed } from '../open-presentation.ts'
import { existingToolFile, existingToolPath } from '../workspace.ts'

/** Create the `univer_import` tool definition. */
export function importTool(ctx: Context, timeoutMs: number) {
  return defineTool({
    name: 'univer_import',
    description:
      'Import an xlsx, csv, tsv, docx, or pptx file as a new Unit inside an explicit draft worktree. While a presentation is already open, do not import another pptx; edit that deck in place.',
    timeoutMs,
    parameters: {
      source: {
        type: 'string',
        required: true,
        description: 'Workspace-relative or absolute Office source path.'
      },
      file: {
        type: 'string',
        required: true,
        description: 'Workspace-relative or absolute target .univer path.'
      },
      worktreeId: {
        type: 'string',
        description: 'Optional branch id; omit to import into the file’s single edit branch.'
      },
      name: { type: 'string', required: true, description: 'Name for the imported Unit.' },
      docType: {
        type: 'string',
        enum: ['traditional', 'modern'],
        description:
          'Doc import mode, .docx only. traditional (default) keeps the source page setup — page size, orientation, margins, and headers/footers survive export. modern converts to a pageless document and rewrites page geometry to a fixed layout, so the source page setup cannot be recovered.'
      }
    },
    output: operationOutput,
    async execute(args, exec) {
      if (args.name.trim().length === 0) {
        throw new UniverError('univer_import requires a non-empty Unit name.', 'INVALID_REQUEST')
      }
      const [target, source] = await Promise.all([
        existingToolFile(exec, args.file),
        existingToolPath(exec, args.source)
      ])
      // The open deck is already that presentation. Importing a pptx again builds a second one.
      assertPptxImportAllowed(ctx, exec, source.path)
      return stripGatewaySuccessEnvelope(
        await ctx.univer.importUnitContent(
          {
            workspace: target.workspace,
            file: target.path,
            sourceWorkspace: source.workspace,
            source: source.path,
            ...(args.worktreeId === undefined ? {} : { worktreeId: worktreeId(args.worktreeId) }),
            name: args.name,
            ...(args.docType === undefined ? {} : { docType: args.docType })
          },
          exec.signal
        )
      )
    },
    presentCall: (args) => ({
      card: 'generic',
      title: operationTitle('import', args.file),
      kind: 'execute'
    })
  })
}
