/**
 * Whether one file address names a presentation this plugin opens in its dedicated panel.
 *
 * Only `.pptx` takes the import path: a `.univer` address may hold any Unit kind, so it keeps the
 * generic preview instead of assuming a Slide Unit exists behind it.
 */
export function isPresentationPath(path: string): boolean {
  return /\.pptx$/i.test(path)
}
