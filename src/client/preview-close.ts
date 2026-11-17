/**
 * Unsaved-change guards for the preview tabs, keyed by session and tab address.
 *
 * A tab body owns whether its file has unsaved changes, while the right-Sidebar close handler runs
 * outside React. This registry is the seam between them: the body registers a guard while mounted,
 * the close handler asks it whether closing would lose work and asks it to show the confirm strip.
 *
 * A guard can be settled before a close that already saved or discarded the branch, because the
 * dirty flag the body exposes lags the Host state by one poll.
 */
export interface UnsavedPreviewGuard {
  /** True while the preview's branch differs from the last saved version. */
  hasUnsavedChanges(): boolean
  /** Ask the mounted preview to show its save/discard/cancel strip. */
  requestClose(): void
}

/** The registration handle: settle it before closing a tab whose branch is already settled. */
export interface UnsavedPreviewRegistration {
  /** Let the next close proceed even before the dirty flag catches up. */
  settle(): void
  /** Remove the guard; a later close no longer consults this preview. */
  dispose(): void
}

interface Entry {
  readonly guard: UnsavedPreviewGuard
  settled: boolean
}

const guards = new Map<string, Entry>()

/** Register the guard for one open preview tab. */
export function registerUnsavedGuard(
  sessionId: string,
  address: string,
  guard: UnsavedPreviewGuard
): UnsavedPreviewRegistration {
  const key = guardKey(sessionId, address)
  const entry: Entry = { guard, settled: false }
  guards.set(key, entry)
  return {
    settle: () => {
      entry.settled = true
    },
    dispose: () => {
      if (guards.get(key) === entry) guards.delete(key)
    }
  }
}

/** The guard for one tab, with its settled flag folded into the answer. */
export function unsavedGuardAt(
  sessionId: string,
  address: string
): { readonly guard: UnsavedPreviewGuard; readonly unsaved: boolean } | undefined {
  const entry = guards.get(guardKey(sessionId, address))
  if (entry === undefined) return undefined
  return { guard: entry.guard, unsaved: !entry.settled && entry.guard.hasUnsavedChanges() }
}

function guardKey(sessionId: string, address: string): string {
  return `${sessionId}\u0000${address}`
}
