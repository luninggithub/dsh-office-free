# PPTX kernel

[简体中文](README.zh-CN.md)

A from-scratch browser kernel for opening and editing `.pptx` files. It does not use Univer or a paid Office SDK. The same page can run on its own, or inside DeepSeek Harness (DSH) as the right-hand preview for `.pptx`.

## Preview

The page can:

- open a `.pptx` and paint slide backgrounds, theme, master, layout, text, pictures, tables, doughnut charts, and pie charts
- edit text on the slide and keep each run's font, size, weight, and color
- move a picture, text box, doughnut chart, or pie chart; resize a picture or either chart; crop a picture
- copy or delete the selected picture, text box, doughnut chart, or pie chart; the crop button's top-left corner sits on that picture's top-right corner
- use Ctrl+C, Ctrl+V, and Ctrl+Z to copy, paste, and undo the selected object
- insert a local PNG, JPEG, GIF, or WebP with **插入图片**, then move, resize, or crop it
- append a blank slide that reuses the current slide's layout
- copy a slide, delete a slide, or drag thumbnails to reorder them; the last remaining slide cannot be deleted
- drag the bar between the thumbnail list and the slide to change both widths
- zoom from 10% to 400%, starting at 100%; **+** and **−** change the scale by 1%. The wheel zooms only while the slide page itself is selected
- scroll the slide horizontally to either edge at any pane width
- overwrite the open file with **保存**, or restore the last saved bytes with **放弃**

Charts other than doughnut and pie charts, SmartArt, animations, transitions, equations, and Office line breaking are not implemented. Text wrapping is approximate.

## Run locally

Node.js `>=22.19.0` and `pnpm@11.23.0`.

```sh
pnpm install
pnpm run pptx:web
node packages/pptx-kernel/web/server.mjs
```

Open `http://127.0.0.1:4173/`. **保存** overwrites a file opened with the file picker, a dropped file handle, or `?open=` from this server. It does not download a new copy. `POST /save?name=` only accepts an existing `.pptx` in `packages/pptx-kernel/web/`.

`web/editor.js` is produced by `pnpm run pptx:web`. Do not edit that bundle by hand. When the page must load a new bundle, bump the `editor.js` query in `packages/pptx-kernel/web/index.html`.

## Install into DSH

From a DeepSeek Harness checkout, add this repository and start the web profile:

```sh
pnpm dsh plugin --profile web add <path-to-this-repo>
pnpm dsh web
```

Restart the Web app after installing. A `.pptx` generated or opened in DSH then opens with this kernel in the session's right-hand preview, the same pane as the built-in viewer. Drag that pane's left edge to change its width. The built-in Office preview remains available from **Open with**.

An authenticated session can also open `http://127.0.0.1:3080/office/`. That route does not accept `/save`. Inside the DSH preview, **保存** still needs a file chosen with **打开** or a handle from the standalone server.

## Development

```sh
pnpm run typecheck
pnpm run test:unit
pnpm run lint
```

Kernel source is `packages/pptx-kernel/src`. The web shell is `packages/pptx-kernel/web`. Kernel tests live in `test/unit/pptx-kernel.test.ts`. The license is Apache-2.0.
