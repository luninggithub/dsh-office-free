import type { JsonValue, UniverUnitKind } from '../service/types.ts'

/** One Unit's state before the agent started changing it. */
export interface EpisodeUnitSnapshot {
  readonly unitId: string
  readonly name: string
  readonly kind: UniverUnitKind
  /** Unit data captured immediately before the agent's first write. */
  readonly data: JsonValue
  /** Revision the branch was seeded from: equal to `capturedRev` when the user had not edited it. */
  readonly seededRev: number
  readonly capturedRev: number
}

/**
 * The branch content to fall back to when the user discards agent changes.
 *
 * Discard normally drops the whole branch, which is the last saved version. A presentation is
 * different: the user may have edited the deck in the preview before the agent touched it, and the
 * requirement is to return to that state rather than to the last save. The snapshot is therefore
 * taken once per Unit, immediately before the agent's first write, and it is the only extra state
 * this plugin keeps for that decision.
 */
export class PresentationEpisodes {
  private readonly snapshots = new Map<string, EpisodeUnitSnapshot>()
  private readonly checked = new Set<string>()

  /** Whether the agent's first write to this Unit already recorded the state to restore. */
  has(file: string, worktreeId: string, unitId: string): boolean {
    const key = keyOf(file, worktreeId, unitId)
    return this.checked.has(key) || this.snapshots.has(key)
  }

  /** Remember that the deck had no preview edits before the agent: a plain discard restores it. */
  mark(file: string, worktreeId: string, unitId: string): void {
    this.checked.add(keyOf(file, worktreeId, unitId))
  }

  record(file: string, worktreeId: string, snapshot: EpisodeUnitSnapshot): void {
    const key = keyOf(file, worktreeId, snapshot.unitId)
    this.checked.add(key)
    this.snapshots.set(key, snapshot)
  }

  /** Read and forget one branch's snapshots; a settled branch cannot be discarded twice. */
  takeWorktree(file: string, worktreeId: string): readonly EpisodeUnitSnapshot[] {
    const prefix = `${file}\u0000${worktreeId}\u0000`
    const taken: EpisodeUnitSnapshot[] = []
    for (const [key, snapshot] of this.snapshots) {
      if (!key.startsWith(prefix)) continue
      this.snapshots.delete(key)
      taken.push(snapshot)
    }
    for (const key of this.checked) {
      if (key.startsWith(prefix)) this.checked.delete(key)
    }
    return taken
  }

  /** Forget a branch after save, discard, or a rebuild that replaced it. */
  forgetWorktree(file: string, worktreeId: string): void {
    this.takeWorktree(file, worktreeId)
  }

  clear(): void {
    this.snapshots.clear()
    this.checked.clear()
  }
}

/**
 * Whether a snapshot carries changes the user made before the agent's first write.
 *
 * The branch copy starts at the trunk revision, so a Unit still at that revision is identical to
 * the saved version. Restoring it would only recreate an identical Unit under a new identity and
 * report a change that does not exist, so Discard settles as a plain branch discard instead.
 */
export function keepsUserEdits(snapshot: EpisodeUnitSnapshot): boolean {
  return snapshot.capturedRev !== snapshot.seededRev
}

function keyOf(file: string, worktreeId: string, unitId: string): string {
  return `${file}\u0000${worktreeId}\u0000${unitId}`
}
