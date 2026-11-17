import type { FileState } from './state.ts'

/**
 * The two decisions the preview offers on unsaved changes.
 *
 * `merge` is Save: it folds the branch into the last saved version. `discard` throws the branch
 * away. There is no separate submission step — a branch is either saved or not.
 */
export type WorktreeReviewAction = 'merge' | 'discard'

/** Result of one save or discard decision. */
export type WorktreeActionResult =
  | {
      readonly ok: true
      readonly action: WorktreeReviewAction
      readonly worktreeId: string
      readonly state: FileState
    }
  | { readonly ok: false; readonly reason: string; readonly state?: FileState }
