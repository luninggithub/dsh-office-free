import * as React from 'react'
import type { BetterSidebarFileViewerProps } from '../better-sidebar.ts'
import { sessionFileAddress } from '../resource-address.ts'
import { UniverPreviewSurface } from './univer-preview-surface.tsx'

/** Viewer id of the on-demand Univer preview registered with dsh-better-sidebar. */
export const UNIVER_FILE_VIEWER_ID = 'univer-office:univer'

/**
 * The `dsh-better-sidebar` file viewer for Univer documents and presentations.
 *
 * This entry point is only reached while that optional plugin is installed: its
 * editor tab type claims file addresses at the same `extension` tier with a
 * longer pattern, so it wins the claim and routes the address through its own
 * file-viewer registry. Hosts without it use the native right-Sidebar tab type
 * in `univer-sidebar-tab.tsx` instead.
 *
 * A `.pptx` is imported into its sibling `.univer` before it renders, so the
 * click and the later conversation agree on one editable Slide Unit. The viewer
 * receives the path rather than the tab address, so it rebuilds the address from
 * the tab's own scope to keep the write-driven reveal pointing at this tab.
 */
export function UniverFileViewer(props: BetterSidebarFileViewerProps): React.ReactElement {
  return (
    <UniverPreviewSurface
      ctx={props.ctx}
      path={props.path}
      sessionId={props.scope.sessionId}
      title={props.title}
      surface="better-sidebar"
      address={sessionFileAddress(props.scope.sessionId, props.scope.cwd, props.path)}
    />
  )
}
