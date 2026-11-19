# PPTX 内核

[English](README.md)

一套从零实现的浏览器 PPTX 内核，用来打开和编辑 `.pptx`。不使用 Univer，也不使用收费的 Office SDK。

当前预览可以：

- 打开 `.pptx`，绘制背景、主题、母版、版式、文字、图片和表格
- 在幻灯片上直接改文字，并保留每一段文字原有的字体、字号、粗细和颜色
- 在当前页后面新增一张空白幻灯片，沿用当前页的版式
- 用「保存」覆盖当前打开的文件
- 用「放弃」回到最近一次保存的内容；还没保存过时，回到打开时的内容

图表、SmartArt、动画、切换、公式，以及和 Office 一致的换行，都还没有实现。文字换行是近似的。

## 预览

```sh
pnpm install
pnpm run pptx:web
node packages/pptx-kernel/web/server.mjs
```

打开 `http://127.0.0.1:4173/`。「保存」会覆盖通过文件选择器、拖入的文件句柄，或本服务 `?open=` 打开的文件，不会另存下载。

## 安装到 DSH

在 DeepSeek Harness 源码目录执行：

```sh
pnpm dsh plugin --profile web add D:\icode\github\dsh-office-free
pnpm dsh web
```

安装并重新打开 Web 后，DSH 里生成或点开的 `.pptx` 会用这个内核显示在会话右侧，和默认预览同一块区域。拖动这一栏左边缘可以左右改变宽度。登录后也可以直接打开 `http://127.0.0.1:3080/office/`。这个地址走 DSH 的登录校验，不提供独立预览服务的 `/save`。

## 开发

需要 Node.js `>=22.19.0` 和 `pnpm@11.23.0`。

```sh
pnpm run typecheck
pnpm run test:unit
pnpm run lint
```

内核源码在 `packages/pptx-kernel`。`web/editor.js` 由 `pnpm run pptx:web` 生成。
