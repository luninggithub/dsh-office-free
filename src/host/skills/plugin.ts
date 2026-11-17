import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import type { Context } from '@deepseek-ai/cordis'
import {
  BUNDLED_SKILL_RANK,
  type SkillCandidate,
  type SkillDefinition,
  type SkillProvider
} from '@deepseek-ai/dsh-skill'

const PROVIDER_NAME = 'univer'
const INVOCATION = { modelInvocable: true, userInvocable: true } as const
const DEFINITIONS = [
  {
    name: 'univer',
    description:
      'Edit and review existing office content the user opened in the Univer side panel: .univer files, presentations (.pptx), documents (.docx), workbooks (.xlsx), Base databases, and Boards. Load this before the matching Unit skill. This plugin edits existing content; it does not create a new Office document, which stays on the deployment default Office path.'
  },
  {
    name: 'univer-sheet',
    description:
      'Read, write, format, calculate, and verify an existing Univer Sheet Unit through DSH tools and the Lite Interface. Use when editing spreadsheet values, formulas, ranges, tables, charts, images, formatting, validation, filters, or pivots in a workbook the user opened, including xlsx/csv/tsv content imported into Univer.'
  },
  {
    name: 'univer-doc',
    description:
      'Read, edit, paginate, chart, inspect, export, and review an existing Univer Doc Unit through DSH tools and the Lite Interface. Use for paragraphs, rich text, lists, tasks, tables, images, charts, headers, footers, and page layout in a document the user opened, including docx content imported into Univer.'
  },
  {
    name: 'univer-slide',
    description:
      'Redesign, edit, inspect, lint, export, and review an existing presentation through Univer Slide Units and the Lite Interface. Use when the user asks to change a deck, page, svg-authored layout, shape, text, image, table, chart, or transition in a presentation already open in the Univer side panel; generated pages should use univer_compile_svg and every changed page should use univer_lint. Do not use it to produce a brand-new deck from a prompt.'
  },
  {
    name: 'univer-base',
    description:
      'Edit, calculate, inspect, export, and review an existing Univer Base database Unit through DSH tools and the Lite Interface. Use for Base tables, fields, records, views, Formula fields, structured references, or Sheet-backed external references in a Base the user opened.'
  },
  {
    name: 'univer-board',
    description:
      'Edit, chart, inspect, and review an existing Univer Board canvas Unit through DSH tools and the Lite Interface. Use for Board shapes, text, connectors, routing, images, native charts, diagrams, and canvas layout in a Board the user opened.'
  },
  {
    name: 'univer-embed',
    description:
      'Embed one Univer Unit inside another through DSH tools and the Lite Interface. Use when a Sheet, Doc, Slide, Base, Board, dashboard, report, presentation, database, or canvas should display or interact with content from another Unit in the same .univer file.'
  },
  {
    name: 'univer-cross-unit-formula',
    description:
      'Author, calculate, update, inspect, and verify cross-Unit formulas through DSH tools and the Lite Interface. Use when a Sheet cell or formula-driven Shape in a Sheet, Doc, Slide, or Board reads a Sheet range or Base table column from another Unit in the same .univer file.'
  }
] as const

const CANDIDATES: readonly SkillCandidate[] = DEFINITIONS.map((definition) => {
  const url = new URL(`../skills/${definition.name}/SKILL.md`, import.meta.url)
  return {
    ...definition,
    invocation: INVOCATION,
    provider: PROVIDER_NAME,
    source: 'bundled',
    resourceBase: {
      kind: 'directory',
      path: fileURLToPath(new URL(`../skills/${definition.name}/`, import.meta.url))
    },
    rank: BUNDLED_SKILL_RANK,
    locator: url
  }
})

const provider: SkillProvider = {
  name: PROVIDER_NAME,
  list: () => Promise.resolve(CANDIDATES),
  async get(candidate): Promise<SkillDefinition> {
    if (!(candidate.locator instanceof URL)) throw new Error('univer skill locator must be a URL')
    return {
      name: candidate.name,
      description: candidate.description,
      invocation: candidate.invocation,
      provider: candidate.provider,
      source: candidate.source,
      ...(candidate.resourceBase === undefined ? {} : { resourceBase: candidate.resourceBase }),
      content: stripFrontmatter(await readFile(candidate.locator, 'utf8'))
    }
  }
}

export const name = 'univer-skills'
export const inject = ['skills']

/**
 * Register version-matched Univer instructions on the DSH skill seam.
 *
 * These skills use their own names and never shadow DSH's bundled Office routes, so creating a
 * new Word, Excel, or PowerPoint document keeps following the deployment default. They carry
 * only the editing workflow for content the user opened in the Univer side panel.
 */
export function apply(ctx: Context): void {
  ctx.skills.registerProvider(() => provider)
}

function stripFrontmatter(value: string): string {
  // Skill files arrive with the checkout's line endings, so match LF and CRLF alike.
  const match = /^---\r?\n[\s\S]*?\r?\n---\r?\n/.exec(value)
  return match === null ? value : value.slice(match[0].length)
}
