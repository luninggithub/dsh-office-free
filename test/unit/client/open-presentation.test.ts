import { describe, expect, it } from 'vitest'
import { registerOpenPreview } from '../../../src/client/preview-registry.ts'
import { keepsOpenPresentation } from '../../../src/client/open-presentation.ts'

const SESSION = 'session-1'

describe('keepsOpenPresentation', () => {
  it('opens a file normally when no presentation tab is mounted', () => {
    expect(keepsOpenPresentation(SESSION, 'D:/decks/other.univer')).toBe(false)
  })

  it('stays on the open deck when a write names that container', () => {
    const dispose = registerOpenPreview({
      sessionId: SESSION,
      container: 'D:/decks/demo.univer',
      address: 'dsh-resource://file/session/session-1/D:/decks/demo.pptx'
    })
    try {
      expect(keepsOpenPresentation(SESSION, 'd:\\decks\\DEMO.univer')).toBe(false)
    } finally {
      dispose()
    }
  })

  it('does not open a second presentation produced while the original deck is open', () => {
    const dispose = registerOpenPreview({
      sessionId: SESSION,
      container: 'D:/decks/demo.univer',
      address: 'dsh-resource://file/session/session-1/D:/decks/demo.pptx'
    })
    try {
      expect(keepsOpenPresentation(SESSION, 'D:/decks/demo-edited.univer')).toBe(true)
      expect(keepsOpenPresentation('other-session', 'D:/decks/demo-edited.univer')).toBe(false)
    } finally {
      dispose()
    }
  })

  it('does not treat a plain .univer preview as a locked presentation', () => {
    const dispose = registerOpenPreview({
      sessionId: SESSION,
      container: 'D:/notes/demo.univer',
      address: 'dsh-resource://file/session/session-1/D:/notes/demo.univer'
    })
    try {
      expect(keepsOpenPresentation(SESSION, 'D:/notes/other.univer')).toBe(false)
    } finally {
      dispose()
    }
  })
})
