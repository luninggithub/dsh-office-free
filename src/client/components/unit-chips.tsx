import * as React from 'react'
import type { TranslateNS } from '@deepseek-ai/dsh-client-ui-slots'
import type { BranchUnit } from '../../shared/wire/state.ts'

const ICONS: Record<Exclude<BranchUnit['kind'], 'unchanged'>, string> = {
  added: '＋',
  modified: '✎',
  deleted: '－',
  conflict: '⚠'
}

/**
 * Navigation chips for the Units the branch shows.
 *
 * The embedded Viewer hides its own file list, so these chips are the only way to switch Units
 * inside the one branch. A chip carries the change mark when the Unit differs from the saved
 * version, and just its name when it does not.
 */
export function UnitChips(props: {
  readonly units: readonly BranchUnit[]
  readonly selected: string | undefined
  readonly t: TranslateNS<'univer'>
  readonly onSelect: (unitId: string) => void
}): React.ReactElement | null {
  if (props.units.length <= 1) return null
  return (
    <div className="uvf_units">
      {props.units.map((unit) => (
        <button
          key={unit.unitId}
          type="button"
          className={`uvf_unit${unit.unitId === props.selected ? ' uvf_unit_on' : ''}`}
          data-kind={unit.kind}
          title={unit.name || props.t(`dock.unit.${unit.kind}`)}
          onClick={() => props.onSelect(unit.unitId)}
        >
          {unit.kind === 'unchanged' ? null : (
            <span className="uvf_unit_icon">{ICONS[unit.kind]}</span>
          )}
          {unit.name || props.t(`dock.unit.${unit.kind}`)}
        </button>
      ))}
    </div>
  )
}

/** Append a selected Unit to an opaque Host Viewer target. */
export function unitViewerUrl(
  url: string | undefined,
  units: readonly BranchUnit[],
  unitId: string | undefined
): string | undefined {
  if (unitId === undefined) return url
  return units.find((entry) => entry.unitId === unitId)?.worktreeUrl ?? url
}
