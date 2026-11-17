/**
 * The file, worktree, and Unit the Univer side panel currently holds for one DSH session.
 *
 * The right panel is the only Univer surface in the browser, so this binding is the browser's
 * editing focus projected back to the Host: a later conversation in the same session continues the
 * exact Unit the user is looking at instead of rebuilding a target. Every identity here is
 * structured; the Client never parses paths or Viewer URLs to recover a target.
 */
export interface OpenPanel {
  /** Authorized source path the user opened, or null when they opened a `.univer` directly. */
  readonly source: string | null
  /** Absolute `.univer` container that owns the displayed Unit. */
  readonly file: string
  readonly worktreeId: string
  readonly unitId: string
  /** Same-origin Viewer target; the Host appends the session scope before projection. */
  readonly viewerUrl: string
  /**
   * The single Slide image the user picked in the Viewer, or null.
   *
   * The Viewer reports it because a picture on a slide is not addressable from the conversation
   * otherwise. It is only a target hint: the Provider re-reads the element and refuses anything that
   * is no longer an image.
   */
  readonly selection: PanelImageSelection | null
}

/** One image element the browser has selected inside the bound Unit. */
export interface PanelImageSelection {
  readonly unitId: string
  readonly pageId: string
  readonly elementId: string
}

/** Result of opening or binding one target in the side panel. */
export type PanelOpenResult =
  | { readonly ok: true; readonly panel: OpenPanel }
  | { readonly ok: false; readonly reason: string }
