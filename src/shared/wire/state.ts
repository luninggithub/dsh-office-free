/** Lifecycle states persisted by the Univer collaboration service. */
export type WorktreeStatus = 'draft' | 'ready' | 'merged' | 'discarded'

/** How one Unit differs from its trunk baseline. */
export type UnitChangeKind = 'added' | 'modified' | 'deleted' | 'conflict'

/** A changed Univer unit read from the Gateway merge preview. */
export interface ChangedUnit {
  readonly unitId: string
  readonly name: string
  readonly type: string | null
  readonly kind: UnitChangeKind
}

/** One Unit the branch shows. */
export interface BranchUnit {
  readonly unitId: string
  readonly name: string
  readonly type: string | null
  /** `unchanged` when this Unit is identical to the last saved version. */
  readonly kind: UnitChangeKind | 'unchanged'
  /** Embedded Viewer target for this Unit inside the branch. */
  readonly worktreeUrl: string
}

/**
 * The one editable branch of a file: work in progress over the last saved version.
 *
 * The browser shows exactly one of these per file, and every edit — a user's keystroke or an agent's
 * write — lands in it. `unsaved` is the only lifecycle distinction left: it is true while the branch
 * differs from the last saved version, which is what enables Save and Discard.
 */
export interface FileBranch {
  readonly worktreeId: string
  readonly name: string
  /** Every Unit the branch shows, not only the changed ones. */
  readonly units: readonly BranchUnit[]
  /** True while this branch differs from the last saved version. */
  readonly unsaved: boolean
  /** Embedded Viewer target for the branch, before a Unit is appended. */
  readonly worktreeUrl: string
}

/** Current collaboration state for one `.univer` file. */
export interface FileState {
  readonly ok: true
  readonly file: string
  readonly gateway: string | null
  readonly gatewayRunning: boolean
  /** The last saved version, as a Viewer target. */
  readonly viewerUrl: string | null
  /** The file's one editable branch, or null before any edit exists. */
  readonly branch: FileBranch | null
}
