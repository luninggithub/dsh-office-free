import { describe, expect, it } from 'vitest'
import {
  keepsUserEdits,
  PresentationEpisodes,
  type EpisodeUnitSnapshot
} from '../../src/host/provider/presentation-episodes.ts'

const snapshot = (overrides: Partial<EpisodeUnitSnapshot> = {}): EpisodeUnitSnapshot => ({
  unitId: 'unit-1',
  name: 'deck',
  kind: 'slide',
  data: { slides: [] },
  seededRev: 3,
  capturedRev: 4,
  ...overrides
})

describe('presentation episodes', () => {
  it('remembers one Unit once so a repeated write does not capture again', () => {
    const episodes = new PresentationEpisodes()
    expect(episodes.has('deck.univer', 'branch', 'unit-1')).toBe(false)

    episodes.record('deck.univer', 'branch', snapshot())
    expect(episodes.has('deck.univer', 'branch', 'unit-1')).toBe(true)
    // A different branch is a different episode, even for the same Unit.
    expect(episodes.has('deck.univer', 'other-branch', 'unit-1')).toBe(false)
  })

  it('returns a branch snapshot exactly once', () => {
    const episodes = new PresentationEpisodes()
    episodes.record('deck.univer', 'branch', snapshot())

    expect(episodes.takeWorktree('deck.univer', 'branch')).toEqual([snapshot()])
    expect(episodes.takeWorktree('deck.univer', 'branch')).toEqual([])
  })

  it('keeps other branches when one branch settles', () => {
    const episodes = new PresentationEpisodes()
    episodes.record('deck.univer', 'branch', snapshot())
    episodes.record('deck.univer', 'other-branch', snapshot({ unitId: 'unit-2' }))

    episodes.forgetWorktree('deck.univer', 'branch')
    expect(episodes.takeWorktree('deck.univer', 'other-branch')).toEqual([
      snapshot({ unitId: 'unit-2' })
    ])
  })

  it('forgets everything on Host unload', () => {
    const episodes = new PresentationEpisodes()
    episodes.record('deck.univer', 'branch', snapshot())
    episodes.mark('deck.univer', 'branch', 'unit-2')

    episodes.clear()
    expect(episodes.has('deck.univer', 'branch', 'unit-1')).toBe(false)
    expect(episodes.has('deck.univer', 'branch', 'unit-2')).toBe(false)
  })
})

describe('pre-agent user edits', () => {
  it('distinguishes a deck the user changed from an untouched one', () => {
    expect(keepsUserEdits(snapshot({ seededRev: 3, capturedRev: 4 }))).toBe(true)
    expect(keepsUserEdits(snapshot({ seededRev: 3, capturedRev: 3 }))).toBe(false)
  })
})
