# PPTX kernel

[简体中文](README.zh-CN.md)

A from-scratch browser kernel for opening and editing `.pptx` files. It does not use Univer or a paid Office SDK.

The current preview can:

- open a `.pptx` and paint slide backgrounds, theme, master, layout, text, pictures, and tables
- edit text in place while keeping each run's font, size, weight, and color
- add a blank slide after the current one, using that slide's layout
- overwrite the open file with **保存**
- revert to the last saved bytes with **放弃**

Charts, SmartArt, animations, transitions, equations, and Office line breaking are not implemented. Text wrapping is approximate.

## Preview

```sh
pnpm install
pnpm run pptx:web
node packages/pptx-kernel/web/server.mjs
```

Open `http://127.0.0.1:4173/`. **保存** overwrites a file opened with the file picker, a dropped file handle, or `?open=` from this server. It does not download a new copy.

## Development

Node.js `>=22.19.0` and `pnpm@11.23.0`.

```sh
pnpm run typecheck
pnpm run test:unit
pnpm run lint
```

The kernel source is `packages/pptx-kernel`. `web/editor.js` is built by `pnpm run pptx:web`.
