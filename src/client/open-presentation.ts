import { sameFile } from './file-key.ts'
import { isPresentationPath } from './presentation-path.ts'
import { openPreviewsOf } from './preview-registry.ts'
import { parseFileAddress } from './resource-address.ts'

/**
 * Whether an agent write must stay on the presentation the user already has open.
 *
 * Image generation and other deck edits name a file, and the preview follows that
 * name. A second `.pptx` — a regenerated deck, an export, or another container —
 * would otherwise become its own tab. While one presentation tab is open, a write
 * that is not that tab's container is ignored: the original tab stays, and the
 * session binding is not moved onto the new file.
 */
export function keepsOpenPresentation(sessionId: string, requestFile: string): boolean {
  const open = openPreviewsOf(sessionId)
  const deckOpen = open.some((preview) => isPresentationAddress(preview.address))
  if (!deckOpen) return false
  return !open.some((preview) => sameFile(preview.container, requestFile))
}

/** Whether a sidebar address is the `.pptx` the user clicked, not its `.univer` cache. */
export function isPresentationAddress(address: string): boolean {
  const parsed = parseFileAddress(address)
  return parsed !== undefined && isPresentationPath(parsed.path)
}
