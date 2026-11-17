import type { Context } from '@deepseek-ai/cordis'
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'
import { describe, expect, it } from 'vitest'
import type { OpenPanel } from '../../src/shared/wire/panel.ts'
import { UniverError } from '../../src/host/service/errors.ts'
import {
  assertPptxExportAllowed,
  assertPptxImportAllowed,
  assertSameOpenPresentation
} from '../../src/host/tools/open-presentation.ts'

const PANEL: OpenPanel = {
  source: 'D:\\decks\\demo.pptx',
  file: 'D:\\decks\\demo.univer',
  worktreeId: 'wt-1',
  unitId: 'unit-1',
  viewerUrl: '/univer-viewer/?file=demo',
  selection: null
}

function ctxOf(panel: OpenPanel | undefined): Context {
  return { univer: { panelOf: () => panel } } as unknown as Context
}

function execOf(sessionId: string | undefined): ToolRunContext {
  if (sessionId === undefined) return {} as ToolRunContext
  return { agent: { session: { header: { id: sessionId } } } } as unknown as ToolRunContext
}

function codeOf(run: () => void): string | undefined {
  try {
    run()
    return undefined
  } catch (error) {
    expect(error).toBeInstanceOf(UniverError)
    return error instanceof UniverError ? error.code : undefined
  }
}

describe('open presentation lock', () => {
  const ctx = ctxOf(PANEL)
  const exec = execOf('session-1')

  it('allows edits of the deck that is already open', () => {
    expect(() => assertSameOpenPresentation(ctx, exec, 'd:/decks/DEMO.univer')).not.toThrow()
    expect(() => assertPptxImportAllowed(ctx, exec, 'D:/sheets/book.xlsx')).not.toThrow()
    expect(() => assertPptxExportAllowed(ctx, exec, 'D:/decks/demo.pdf')).not.toThrow()
  })

  it('rejects a second presentation while the original deck is open', () => {
    expect(codeOf(() => assertSameOpenPresentation(ctx, exec, 'D:/decks/other.univer'))).toBe(
      'PRESENTATION_ALREADY_OPEN'
    )
    expect(codeOf(() => assertPptxImportAllowed(ctx, exec, 'D:/decks/other.pptx'))).toBe(
      'PRESENTATION_ALREADY_OPEN'
    )
    expect(codeOf(() => assertPptxExportAllowed(ctx, exec, 'D:/decks/demo.pptx'))).toBe(
      'PRESENTATION_ALREADY_OPEN'
    )
    expect(codeOf(() => assertPptxExportAllowed(ctx, exec, 'D:/decks/demo-new.pptx'))).toBe(
      'PRESENTATION_ALREADY_OPEN'
    )
  })

  it('does nothing when the session has no open presentation', () => {
    const closed = ctxOf(undefined)
    expect(() => assertSameOpenPresentation(closed, exec, 'D:/decks/other.univer')).not.toThrow()
    expect(() => assertPptxImportAllowed(closed, exec, 'D:/decks/other.pptx')).not.toThrow()
    expect(() =>
      assertPptxExportAllowed(closed, execOf(undefined), 'D:/decks/new.pptx')
    ).not.toThrow()
  })
})
