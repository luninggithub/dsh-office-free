import type { Context } from '@deepseek-ai/cordis'
import type { ResolvedConfig } from '../config.ts'
import type {} from '../service/univer-service.ts'
import { apiTool } from './definitions/api.ts'
import { compileSvgTool } from './definitions/compile-svg.ts'
import { executeTool } from './definitions/execute.ts'
import { exportTool } from './definitions/export.ts'
import { imageEditTool } from './definitions/image-edit.ts'
import { importTool } from './definitions/import.ts'
import { inspectTool } from './definitions/inspect.ts'
import { lintTool } from './definitions/lint.ts'
import { printPdfTool } from './definitions/print-pdf.ts'
import { newTool } from './definitions/new.ts'
import { resourcesTool } from './definitions/resources.ts'
import { screenshotTool } from './definitions/screenshot.ts'
import { statusTool } from './definitions/status.ts'
import { unitTool } from './definitions/unit.ts'
import { worktreeTool } from './definitions/worktree.ts'
import { withUniverErrorContent } from './presentation.ts'

export const inject = ['univer', 'tools']
export const name = 'univer-tools'

/** Register model-facing domain tools over `ctx.univer`. */
export function apply(ctx: Context, config: ResolvedConfig): void {
  const gatewayReadTimeoutMs = config.gatewayStartupTimeoutMs + config.gatewayRequestTimeoutMs
  const gatewayWriteTimeoutMs = config.gatewayStartupTimeoutMs + config.gatewayMutationTimeoutMs
  const unitContentTimeoutMs = config.gatewayStartupTimeoutMs + config.unitContentOperationTimeoutMs
  const screenshotTimeoutMs = config.gatewayStartupTimeoutMs + config.screenshotOperationTimeoutMs
  const printPdfTimeoutMs = config.gatewayStartupTimeoutMs + config.printPdfOperationTimeoutMs
  const imageEditTimeoutMs = config.gatewayStartupTimeoutMs + config.imageOperationTimeoutMs
  ctx.tools.register(withUniverErrorContent(newTool(ctx, gatewayWriteTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(statusTool(ctx, gatewayReadTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(worktreeTool(ctx, gatewayWriteTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(unitTool(ctx, gatewayWriteTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(importTool(ctx, unitContentTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(inspectTool(ctx, unitContentTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(executeTool(ctx, unitContentTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(exportTool(ctx, unitContentTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(lintTool(ctx, unitContentTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(printPdfTool(ctx, printPdfTimeoutMs)))
  ctx.tools.register(withUniverErrorContent(compileSvgTool(ctx, unitContentTimeoutMs)))
  // Image editing needs an Ark API key, but the tool stays registered either way: an install
  // without one must answer "configure imageApiKey or ARK_API_KEY" instead of advertising no tool
  // at all. A missing tool is invisible to the model, which then falls back to an unrelated
  // workflow and the user sees an explanation about the tool not existing rather than a fixable
  // configuration gap. The Provider owns the stable `IMAGE_API_KEY_MISSING` refusal.
  ctx.tools.register(withUniverErrorContent(imageEditTool(ctx, imageEditTimeoutMs)))
  // A screenshot result must durably reference image bytes; advertise the tool only while
  // the deployment has an attachment store, and keep the execution-time re-check defensive.
  ctx.inject(['attachments'], (imageCtx) => {
    imageCtx.tools.register(withUniverErrorContent(screenshotTool(imageCtx, screenshotTimeoutMs)))
  })
  ctx.tools.register(withUniverErrorContent(apiTool(ctx)))
  ctx.tools.register(withUniverErrorContent(resourcesTool(ctx, config.resourceOperationTimeoutMs)))
  ctx.on('tools/pre-execute', (exec, next) => {
    if (exec.name !== 'univer_worktree' || !isRecord(exec.arguments)) return next()
    const action = exec.arguments.action
    if (action !== 'merge' && action !== 'discard') return next()
    return Promise.resolve({
      kind: 'ask',
      reason:
        action === 'merge'
          ? 'Saving publishes the file’s edit branch as the current version.'
          : 'Discarding permanently removes the file’s unsaved changes.'
    })
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
