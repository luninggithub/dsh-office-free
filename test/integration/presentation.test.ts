import { readdir, readFile, stat, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { readZip } from '../support/ooxml.ts'
import { startGateway, type GatewayHarness } from '../support/gateway.ts'

const insertText = (text: string): string =>
  `const slide = presentation.getSlideByIndex(0);` +
  `const shape = slide.insertShape({ shapeType: api.Enum.ShapeTypeEnum.RoundRect, transform: { left: 40, top: 40, width: 320, height: 80 } });` +
  `if (!shape) throw new Error("Cannot insert Slide shape");` +
  `shape.getText().setText(${JSON.stringify(text)}); return shape.getId();`

describe('presentation save', () => {
  let harness: GatewayHarness | undefined

  afterEach(async () => {
    await harness?.dispose()
    harness = undefined
  })

  it('writes a saved deck back to the .pptx it was opened from and keeps the previous version', async () => {
    harness = await startGateway({ fileName: 'deck.univer' })
    const deck = join(harness.workspace, 'deck.pptx')
    await writeFile(deck, 'previous deck')

    const worktreeId = await harness.createWorktree('presentation save')
    const unitId = await harness.createUnit(worktreeId, 'slide', 'deck')
    const written = await harness.execute(worktreeId, unitId, insertText('Saved page'))
    expect(written.committed).toBe(true)

    const saved = await harness.service.worktreeAction({
      ...harness.scoped,
      worktreeId,
      action: 'merge'
    })
    expect(saved.ok, JSON.stringify(saved)).toBe(true)
    // The branch was merged, so the toolbar has nothing left to review.
    expect(saved.state?.branch).toBeNull()

    // The container is only a cache: the saved deck itself is a real OOXML package now.
    expect(readZip(await readFile(deck)).has('ppt/presentation.xml')).toBe(true)
    expect(await readFile(`${deck}.bak`, 'utf8')).toBe('previous deck')

    // No staged export is left beside the deck.
    const leftovers = (await readdir(harness.workspace)).filter((name) => name.includes('.dsh-'))
    expect(leftovers).toEqual([])

    // Trunk reports exactly the deck that was just written, so saving does not duplicate a Unit.
    const status = await harness.service.status(harness.scoped)
    const trunkSlides = status.result.trunk.units.filter(
      (unit: { type: number }) => unit.type === 3
    )
    expect(trunkSlides.length).toBe(1)

    // The container is touched after the deck was written: opening the file again must not treat
    // the deck as an external change and re-import it.
    expect((await stat(harness.file)).mtimeMs).toBeGreaterThanOrEqual((await stat(deck)).mtimeMs)
  })

  it('keeps edits made before the agent touched the deck when the branch is discarded', async () => {
    harness = await startGateway({ fileName: 'deck.univer' })
    const deck = join(harness.workspace, 'deck.pptx')
    const worktreeId = await harness.createWorktree('presentation discard')
    const unitId = await harness.createUnit(worktreeId, 'slide', 'deck')

    // The user edits the preview before the deck is known as a presentation: this revision is the
    // state Discard must return to.
    const userEdit = await harness.execute(worktreeId, unitId, insertText('USER KEEP'))
    expect(userEdit.committed).toBe(true)
    await writeFile(deck, 'previous deck')

    const agentEdit = await harness.execute(worktreeId, unitId, insertText('AGENT DROP'))
    expect(agentEdit.committed).toBe(true)

    const discarded = await harness.service.worktreeAction({
      ...harness.scoped,
      worktreeId,
      action: 'discard'
    })
    expect(discarded.ok, JSON.stringify(discarded)).toBe(true)

    // Discarding the agent's branch restores the pre-agent state as a new unsaved branch, because
    // otherwise the user's own edits would be lost with it.
    const branch = discarded.state?.branch
    expect(branch).not.toBeNull()
    expect(branch.unsaved).toBe(true)
    expect(branch.worktreeId).not.toBe(worktreeId)

    const restored = branch.units[0]
    const inspection = await harness.service.inspectUnitContent({
      workspace: harness.workspace,
      file: harness.file,
      worktreeId: branch.worktreeId,
      unitId: restored.unitId
    })
    const content = JSON.stringify(inspection.result)
    expect(content).toContain('USER KEEP')
    expect(content).not.toContain('AGENT DROP')
  })
})
