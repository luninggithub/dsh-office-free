---
name: univer
description: Edit and review existing office content the user opened in the Univer side panel: .univer files, presentations (.pptx), documents (.docx), workbooks (.xlsx), Base databases, and Boards. Load this before the matching Unit skill. This plugin edits existing content; it does not create a new Office document, which stays on the deployment default Office path.
---

# Univer files

Use the structured `univer_*` tools whenever the task changes, reads, converts, or reviews office content the user already has. Do not wait for the user to name a tool. Do not invoke a global `univer` CLI, edit `.univer` storage directly, or substitute openpyxl, python-pptx, python-docx, ZIP manipulation, or another writer.

## Edit, do not generate

This plugin owns modification of existing content, not generation:

- A request to create a brand-new Word, Excel, or PowerPoint document from a prompt is not this skill's task. Leave it to the deployment's default Office path; do not route it here.
- A request to change a file the user opened in the Univer side panel is this skill's task. The panel already imported the source into a `.univer` draft, so continue that Unit instead of rebuilding it.
- `univer_status` reports an `openPanel` object while the right-side Univer panel holds a file. Treat that file, worktree, and Unit as the active target for the conversation and edit it rather than producing a second copy.

## Start immediately

- Existing `.univer`: call `univer_status` before selecting a Unit or worktree.
- A file already open in the right-side Univer panel: read `openPanel` from `univer_status` and continue that exact `file`, `worktreeId`, and `unitId`.
- New empty `.univer` container: call `univer_new`, then `univer_worktree` with `action: "create"`.
- A source the user attached but did not open (`.xlsx`, `.csv`, `.tsv`, `.docx`, `.pptx`): create the target `.univer` and draft worktree, then call `univer_import`.
- Before authoring content, load the matching Unit skill: `univer-sheet`, `univer-doc`, `univer-slide`, `univer-base`, or `univer-board`.
- For an Embed or a Chart backed by another Unit's data, also load `univer-embed`. For formulas that read another Unit, also load `univer-cross-unit-formula`.

## Mental model

- A `.univer` file is the authoritative multi-Unit container. Each Sheet, Doc, Slide, Base, or Board is a top-level Unit with a stable `unitId`.
- Sheet names, pages, paragraphs, tables, ranges, shapes, fields, records, and views live inside a Unit; none substitutes for `unitId`.
- `trunk` is the last saved version. A file has at most one branch of unsaved changes, reported as `branch` by `univer_status`, and every edit — the agent's and the user's — lands there.
- `worktreeId` is optional on every edit tool. Omit it and the Host uses the file's only branch; pass the `branch.worktreeId` from `univer_status` only when you need to name it explicitly.
- Every content write requires `file` and `unitId`; the branch is resolved for you.
- `univer_execute` persists only when Facade mutations occurred. A read-only execution produces no revision.
- Saving (`merge`) and discarding (`discard`) are the user's decisions in the preview, behind explicit approval. Do not save or discard to finish a task.
- A presentation opened as a `.pptx` keeps that file as its identity: saving writes the deck back to the original `.pptx` and keeps the previous version as `<name>.pptx.bak`; discarding reverts only your edits and keeps pages the user changed in the preview themselves.
- Tool success is not correctness evidence. Read the changed model back and verify task-specific assertions.

## SVG resources

When a Unit needs icons, logos, emoji, or illustrations, use the bundled resource library instead of inventing placeholders:

1. Call `univer_resources` with `action: "registries"` when registry choice matters.
2. Call `univer_resources` with `action: "find"` and one or more semantic `queries`; retain the exact returned handles.
3. Call it with `action: "export"`, those `handles`, and an explicit workspace `output` directory.
4. Reference the exported SVG files from Slide SVG, or read one handle with `action: "read"` when inline SVG text is required.

Only `colorEditable: true` resources may follow an authored color. Fixed logos, color emoji, and illustrations keep their intrinsic colors. Reuse the exported file; do not copy partial path data from it.

## Facade API lookup

- No relevant class or API label is known: use `find` with API-name keywords or identifier fragments.
- A class is known: use `show` on the class to inspect its APIs.
- A type or exact `Class.member` API label is known: use `show` on that label for signatures, documentation, and examples.

`find` is case-insensitive. Each query is searched independently and returns its own matches; queries are not combined as AND, and `find` does not interpret intent. Pass a useful returned label to `show` instead of searching that label again.

## Required workflow

1. Call `univer_status` to discover Unit IDs and the file's branch.
2. Ensure the branch exists: call `univer_worktree` with `action: "create"` only when `univer_status` reports no branch. There is nothing to submit or reopen.
3. Create a Unit with `univer_unit`, or import one with `univer_import`.
4. Load the matching Unit skill before writing Facade code.
5. Resolve unfamiliar Facade usage with `univer_api` following the lookup rules above. Never guess an unfamiliar signature, parameter type, or enum.
6. Mutate through `univer_execute`, or through `univer_compile_svg` for generated Slide page content.
7. Read the changed scope with `univer_inspect`; use a fresh read-only `univer_execute` when inspection omits a required model field.
8. For every changed Slide page, call `univer_lint` and resolve or explicitly justify each finding.
9. For visually relevant changes, call `univer_screenshot` with an explicit workspace output directory and the narrowest useful Unit-specific target. Inspect every returned image; screenshots complement rather than replace structural readback or Slide lint.
10. Export with `univer_export`, or print a PDF with `univer_print_pdf`, only when requested and only from the verified scope.
11. Confirm the final state with `univer_status`, then report. Saving and discarding stay with the user in the preview toolbar.

`univer_screenshot` returns model-visible PNG evidence and writes the same PNGs under its explicit `output` directory. It requires an image-capable current model route. The DSH client renders the live branch content and the Save / Discard toolbar in that same right-side preview; it remains the user-facing handoff surface because there is no model-facing show/open tool. Unsaved changes stay on the branch until the user clicks Save or Discard, and they survive neither a tab close nor an app restart, so never rely on them as a saved result. The user can also edit the page directly, so re-read the changed scope before a write when the user reports a manual change. Claim visual verification only for images actually inspected in the current result.

`merge` and `discard` only when the user explicitly requests that outcome. Both are approval-gated review decisions, not routine completion steps.

## Rework after feedback

Continue in the same branch: revisions the user requested belong to the same task, and the branch keeps accumulating them.

1. Call `univer_status` and read `branch`; when the client shows unsaved changes, read the changed scope first because the user may have edited it by hand.
2. Make the remaining changes and repeat the complete readback and Unit-specific verification.
3. Report the result. Saving is the user's click, not yours.

## Tool map

| Stage | Tool | Use |
| --- | --- | --- |
| Start | `univer_new` | Create an empty `.univer`; never overwrites and never creates an implicit Unit. |
| Start | `univer_status` | List trunk Units and the file's branch, or inspect one explicit scope. |
| Start | `univer_worktree` | `create` ensures the one branch; `merge` saves it, writing a presentation back to its original `.pptx`; `discard` throws it away. |
| Start | `univer_unit` | Create or remove a Sheet, Doc, Slide, Base, or Board in a draft worktree. |
| Start | `univer_import` | Import local xlsx, csv, tsv, docx, or pptx as a new Unit. |
| Write | `univer_execute` | Run version-matched Facade JavaScript against one Unit in a draft worktree. |
| Write | `univer_compile_svg` | Compile workspace SVG into one explicit Slide page with browser text metrics. |
| Verify | `univer_inspect` | Read structured Unit content from trunk or one worktree. |
| Verify | `univer_lint` | Check Slide text off-page, container escape, and text overlap. |
| Verify | `univer_screenshot` | Render Sheet, Doc, Slide, Base, or Board PNG evidence and return it to an image-capable model. |
| Reference | `univer_api` | Find an unknown name, or show a known class, API, or type. |
| Reference | `univer_resources` | List/find/read/export bundled SVG resources or clear their download cache. |
| Deliver | `univer_export` / `univer_print_pdf` | Export Office files or print Sheet, Doc, Slide, or Board to PDF. |

## Facade execution

`univer_execute` injects `univerAPI`, `api` (the same object), and one Unit-specific handle:

- Sheet: `workbook`
- Doc: `doc`
- Slide: `presentation`
- Base: `base`
- Board: `board`

Do not redeclare injected variables. Use `code` only for a small snippet. For multi-line or reusable Facade logic, write a workspace JavaScript body file and pass it as `codeFile`; provide exactly one of `code` or `codeFile`.

Execution is Node ESM: `require` is unavailable, while dynamic `import()` and `Buffer` are available. For an explicit local workspace image, generate its data URI inside the execution instead of putting Base64 in the conversation:

```js
const { readFile } = await import("node:fs/promises");
const bytes = await readFile("/absolute/session/workspace/image.png");
const imageDataUri = `data:image/png;base64,${Buffer.from(bytes).toString("base64")}`;
```

Read only task assets inside the session workspace. Do not print or return the data URI. Resolve exact Facade methods with `univer_api`, check boolean/null returns, and retain stable IDs needed by later operations.

## Import and export

- Import accepts workspace files only; URL import is unavailable.
- Export Sheet or Base to `.xlsx`, `.csv`, or `.tsv`; Doc to `.docx`; Slide to `.pptx`.
- Board export is unsupported.
- Export uses an explicit Unit and an explicit trunk or worktree scope. Recalculate formulas and finish readback before exporting.

## Unsupported CLI-only capabilities

Do not invent equivalents for CLI maintenance, daemon/configuration, `compile-typst`, optimization, or shell command help. Use the bundled DSH tools for resources and screenshots; if the task requires another missing capability, report that exact gap.

## Failure recovery

Tool failures use `Error [CODE]: message`. Route recovery by the code, not free text.

- For `GATEWAY_UNAVAILABLE` or `GATEWAY_REQUEST_TIMEOUT`, retry one `univer_status` read. After a write timeout, inspect status and content before continuing.
- For `FILE_PERMISSION_DENIED` or `SESSION_SCOPE_DENIED`, use an accessible in-workspace path or ask the user to correct access; do not retry the same path.
- For worktree or Unit state errors, refresh with `univer_status`. A branch left `merged` or `discarded` by an earlier Turn is terminal at the Gateway layer; `univer_worktree` with `action: "create"` ensures a fresh writable one.
