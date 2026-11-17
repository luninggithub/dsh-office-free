import { randomUUID } from 'node:crypto'
import { copyFile, rm, stat, utimes } from 'node:fs/promises'
import { basename, dirname, extname, join } from 'node:path'

/**
 * The presentation a `.univer` container caches.
 *
 * A deck the browser opens is imported into a sibling container named after it, so the source is
 * recoverable from the container alone. That derivation is what lets a later save write the deck
 * back without asking the model or the browser which `.pptx` it came from.
 */
export function presentationSiblingOf(container: string): string {
  const name = basename(container)
  const extension = extname(name)
  return join(dirname(container), `${name.slice(0, name.length - extension.length)}.pptx`)
}

/** The existing presentation beside one container, or undefined when it is not a deck cache. */
export async function presentationSiblingExists(container: string): Promise<string | undefined> {
  const candidate = presentationSiblingOf(container)
  try {
    const found = await stat(candidate)
    return found.isFile() ? candidate : undefined
  } catch (error) {
    // Only a missing sibling means "not a presentation"; a permission failure must surface.
    if (nodeErrorCode(error) === 'ENOENT') return undefined
    throw error
  }
}

/** Whether the presentation changed after the container was last written. */
export async function presentationChanged(source: string, container: string): Promise<boolean> {
  try {
    const [from, into] = await Promise.all([stat(source), stat(container)])
    return from.mtimeMs > into.mtimeMs
  } catch (error) {
    if (nodeErrorCode(error) === 'ENOENT') return false
    throw error
  }
}

/**
 * Where one save stages its exported deck.
 *
 * The path stays beside the source and keeps the `.pptx` extension: the export worker picks the
 * format from it, and a same-directory copy keeps the final replacement on one volume and inside
 * the workspace that authorized the write.
 */
export function stagedPresentationPath(source: string): string {
  const name = basename(source)
  const extension = extname(name)
  const stem = name.slice(0, name.length - extension.length)
  return join(dirname(source), `${stem}.dsh-${randomUUID().slice(0, 8)}.tmp${extension}`)
}

/**
 * Replace an original presentation with a staged export, keeping the previous version.
 *
 * The previous file is copied to `<source>.bak` first so an interrupted replacement is always
 * recoverable, and an export that cannot finish leaves the original presentation untouched.
 */
export async function commitStagedPresentation(staged: string, source: string): Promise<void> {
  try {
    await copyFile(source, `${source}.bak`)
  } catch (error) {
    // A first save has nothing to preserve; every other failure must abort the replacement.
    if (nodeErrorCode(error) !== 'ENOENT') throw error
  }
  await copyFile(staged, source)
}

/**
 * Record that the container now matches the presentation it was saved from.
 *
 * Staleness is mtime-based, and the replacement above writes the presentation after the merge, so
 * without this touch the next open would re-import an unchanged deck. A failed touch therefore
 * costs one redundant import and must not fail a save that already persisted the user's work.
 */
export async function markPresentationImported(container: string): Promise<void> {
  const now = new Date()
  await utimes(container, now, now).catch(() => {
    // The deck is saved; only the provenance timestamp is stale.
  })
}

/** Remove a staged export; failure is expected when the export never produced a file. */
export async function discardStagedPresentation(staged: string): Promise<void> {
  await rm(staged, { force: true })
}

function nodeErrorCode(error: unknown): string | undefined {
  if (typeof error !== 'object' || error === null || !('code' in error)) return undefined
  return typeof error.code === 'string' ? error.code : undefined
}
