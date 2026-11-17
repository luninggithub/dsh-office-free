import { mkdir, mkdtemp, readFile, rm, stat, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  commitStagedPresentation,
  discardStagedPresentation,
  markPresentationImported,
  presentationChanged,
  presentationSiblingExists,
  presentationSiblingOf,
  stagedPresentationPath
} from '../../src/host/provider/presentation-files.ts'

let workspace: string

beforeEach(async () => {
  workspace = await mkdtemp(join(tmpdir(), 'dsh-presentation-files-'))
})

afterEach(async () => {
  await rm(workspace, { recursive: true, force: true })
})

describe('presentation sibling', () => {
  it('derives the deck a container caches, in the container directories', () => {
    expect(presentationSiblingOf(join(workspace, 'deck.univer'))).toBe(join(workspace, 'deck.pptx'))
    expect(presentationSiblingOf(join(workspace, 'quarterly.report.univer'))).toBe(
      join(workspace, 'quarterly.report.pptx')
    )
  })

  it('answers only an existing file beside the container', async () => {
    const container = join(workspace, 'deck.univer')
    expect(await presentationSiblingExists(container)).toBeUndefined()

    await writeFile(presentationSiblingOf(container), 'deck')
    expect(await presentationSiblingExists(container)).toBe(presentationSiblingOf(container))
  })

  it('treats a container directory as not being a deck', async () => {
    const container = join(workspace, 'folder.univer')
    const folder = join(workspace, 'folder.pptx')
    await mkdir(folder)
    expect(await presentationSiblingExists(container)).toBeUndefined()
  })
})

describe('presentation staleness', () => {
  it('reports a deck written after its container as changed', async () => {
    const container = join(workspace, 'deck.univer')
    const source = join(workspace, 'deck.pptx')
    await writeFile(container, 'container')
    await writeFile(source, 'saved deck')

    const older = new Date(Date.now() - 60_000)
    await utimes(container, older, older)
    expect(await presentationChanged(source, container)).toBe(true)

    // Saving writes the deck first and then aligns the container, so a saved pair must stay quiet.
    await markPresentationImported(container)
    expect(await presentationChanged(source, container)).toBe(false)
  })

  it('reports a missing pair as unchanged instead of failing', async () => {
    expect(
      await presentationChanged(join(workspace, 'absent.pptx'), join(workspace, 'absent.univer'))
    ).toBe(false)
  })
})

describe('staged replacement', () => {
  it('keeps the previous deck as a backup and replaces it with the export', async () => {
    const source = join(workspace, 'deck.pptx')
    const staged = stagedPresentationPath(source)
    await writeFile(source, 'previous deck')
    await writeFile(staged, 'exported deck')

    // The staged file stays beside the source and keeps the `.pptx` extension so the export worker
    // picks the format from it.
    expect(dirname(staged)).toBe(workspace)
    expect(staged.endsWith('.tmp.pptx')).toBe(true)

    await commitStagedPresentation(staged, source)
    expect(await readFile(source, 'utf8')).toBe('exported deck')
    expect(await readFile(`${source}.bak`, 'utf8')).toBe('previous deck')
  })

  it('does not require a previous deck for a first save', async () => {
    const source = join(workspace, 'fresh.pptx')
    const staged = stagedPresentationPath(source)
    await writeFile(staged, 'first deck')

    await commitStagedPresentation(staged, source)
    expect(await readFile(source, 'utf8')).toBe('first deck')
  })

  it('removes a staged export, tolerating one that was never written', async () => {
    const staged = stagedPresentationPath(join(workspace, 'deck.pptx'))
    await writeFile(staged, 'exported deck')
    await discardStagedPresentation(staged)
    await expect(stat(staged)).rejects.toMatchObject({ code: 'ENOENT' })

    await discardStagedPresentation(staged)
  })
})
