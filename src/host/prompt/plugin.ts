import type { Context } from '@deepseek-ai/cordis'

export const name = 'univer-prompt'
export const inject = ['systemPrompt']

/**
 * Placement between the deliverable file-reference guidance and structured output.
 * External contributions may use any finite order; 9500 keeps this block out of the
 * first-party ranges that own the opener and the environment suffix.
 */
const SECTION_ORDER = 9500

const TEXT = `Univer editing focus:
- The right Univer panel is the only place this plugin renders content. When the DSH client holds a target there, \`univer_status\` reports it as \`openPanel\`. Treat that file, worktreeId, and unitId as the active target and edit it with the \`univer_*\` tools plus the \`univer\` and Unit skills instead of producing another copy.
- A presentation the user opened as a \`.pptx\` keeps that deck as the target: edit its Slide Unit on top of the existing pages. Saving writes the deck back to the original \`.pptx\` and keeps the previous version beside it as a \`.bak\`; discarding reverts your edits while keeping pages the user changed in the panel themselves.
- Panel actions stay user-driven: the user compares, saves, and discards from the panel toolbar. Never merge or discard a worktree on your own.
- Image requests about a picture on the open presentation are a direct model call: call \`univer_image_edit\` and let the panel supply the file, Unit, page, and element. The generated image is written back onto that same deck. Do not create, import, or export another \`.pptx\`, and do not hand the job to the Office path or python-pptx. A missing selection is a reason to ask the user to select the image, not a reason to generate a new file.
- This plugin edits existing content only. Creating a new Word, Excel, or PowerPoint document is not routed here; the deployment default Office path owns that.`

/**
 * Tell the model how an open Univer panel changes task routing.
 *
 * The section is static: which file is open is browser state, so the concrete target reaches the
 * model only through the structured `univer_status` result rather than through prompt text.
 */
export function apply(ctx: Context): void {
  ctx.inject(['systemPrompt'], (promptCtx: Context) => {
    promptCtx.effect(
      () =>
        promptCtx.systemPrompt.section({
          name: 'univer:panel-focus',
          order: SECTION_ORDER,
          text: TEXT
        }),
      'univer: panel routing'
    )
  })
}
