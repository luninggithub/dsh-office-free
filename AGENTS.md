# AGENTS.md

This file defines repository-specific instructions for the from-scratch PPTX kernel in `packages/pptx-kernel`. Describe only the current implementation.

## 1. Commands and validation

Use Node.js `>=22.19.0` and `pnpm@11.23.0`.

```sh
pnpm install
pnpm run lint
pnpm run typecheck
pnpm run test:unit
pnpm run pptx:web
node packages/pptx-kernel/web/server.mjs
```

- Treat `oxlint.config.ts` as project policy. Fix lint errors instead of weakening rules.
- `lint:fix` and `format` mutate files. Run them only when their full write scope is intended.
- `format` and `format:check` are repository-wide. For a narrow change, check the touched files with `pnpm exec oxfmt --check <files>`.
- `pnpm run pptx:web` rebuilds `packages/pptx-kernel/web/editor.js`. Do not hand-edit that bundle. Bump the `editor.js` query in `index.html` when the page must load the new bundle.
- The preview server serves `packages/pptx-kernel/web` on `127.0.0.1:4173` and accepts `POST /save?name=` only for an existing `.pptx` in that directory.
- Run `pnpm run typecheck` for TypeScript changes and `git diff --check` before every handoff. Report the exact commands that completed successfully.

## 2. Kernel boundaries

- `@dsh/pptx-kernel` has no Univer or paid Office SDK dependency. Do not add one.
- The package is isomorphic: browser and Node 22. `src/` must not import Node-only modules. ZIP uses `DecompressionStream` / `CompressionStream`.
- Web UI code lives in `packages/pptx-kernel/web/`. It imports the kernel through `../src/index.ts`.
- Content writes stay inside the open PPTX package. Untouched ZIP parts are copied through on save.
- `save()` overwrites the open file. The page does not download a new file.
- A new slide is blank, appended at the end, and reuses the current slide's layout.
- Thumbnail order, copy, and delete rewrite the slide list in the package. The last remaining slide cannot be deleted.
- Text edits keep existing run formatting. Do not put a whole shape's text into the first run.
- Picture and text-box moves, picture resize and crop, and shape deletion rewrite that shape in the open package.
- The preview zoom is an explicit scale from 10% to 400%, default 100%. The wheel changes that scale only while the slide page itself is selected; otherwise it scrolls the viewport.

## 3. TypeScript

- Use ESM, explicit `.ts` relative imports, and top-level `import type` declarations. Keep `strict`, `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`, and `verbatimModuleSyntax` passing.
- Do not bypass type safety with broad assertions or unexplained `any`.
- Comments explain non-obvious behavior, failure modes, timing, ownership, or security constraints.

## 4. Testing

- Test observable behavior. Place kernel tests in `test/unit/pptx-kernel.test.ts`.
- A bug fix lands with a regression test.
- Do not commit user `.pptx` files. `packages/pptx-kernel/.gitignore` ignores `web/*.pptx`.
