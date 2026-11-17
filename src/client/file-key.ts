/**
 * Identity of one local file path inside the browser.
 *
 * The same file reaches the Client as a tool argument, a Host projection, and a
 * Sidebar address, and those three spell separators and (on Windows) letter case
 * differently. Everything that keys a preview by file — the mounted-preview
 * registry, the Host binding comparison — goes through here so one file is one
 * key.
 */

/**
 * Canonical key for one path: `/` separators, and folded case on Windows.
 *
 * POSIX paths stay case-sensitive: `A.univer` and `a.univer` are different
 * files there, while Windows cannot tell them apart.
 */
export function fileKey(path: string): string {
  const slashed = path.replace(/\\/g, '/')
  const windows = slashed.startsWith('//') || /^[A-Za-z]:\//.test(slashed)
  return windows ? slashed.toLowerCase() : slashed
}

/** Whether two spellings name the same file. */
export function sameFile(left: string, right: string): boolean {
  return fileKey(left) === fileKey(right)
}
