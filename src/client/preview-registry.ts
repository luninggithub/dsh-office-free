import { fileKey } from './file-key.ts'

/**
 * The preview the user currently has open, as this plugin's own record.
 *
 * A DSH right-Sidebar tab is addressed by resource address, and the address a
 * `.pptx` opens under is the deck, not the `.univer` container the agent writes.
 * An agent write therefore cannot name the user's tab by itself; it has to look
 * up which preview already shows that container. This registry is that lookup:
 * each mounted preview records its own address, so a later write reveals the
 * exact open tab instead of opening a second one.
 *
 * Entries are keyed by session and container path, and they exist only while the
 * preview is mounted — a closed tab reopens under the container's own address.
 */

/** One mounted preview. */
export interface OpenPreview {
  readonly sessionId: string
  /** Absolute `.univer` container the preview shows. */
  readonly container: string
  /** The address that tab was opened under, which is also its identity. */
  readonly address: string
}

const previews = new Map<string, OpenPreview>()

const keyOf = (sessionId: string, container: string): string =>
  `${sessionId}\u0000${fileKey(container)}`

/**
 * Record one mounted preview.
 * @returns the disposer that forgets it on unmount.
 */
export function registerOpenPreview(entry: OpenPreview): () => void {
  const key = keyOf(entry.sessionId, entry.container)
  previews.set(key, entry)
  return () => {
    if (previews.get(key) === entry) previews.delete(key)
  }
}

/** The address of the open preview for one container, when the user has it open. */
export function addressOfOpenPreview(sessionId: string, container: string): string | undefined {
  return previews.get(keyOf(sessionId, container))?.address
}
