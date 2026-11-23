# PPTX 内核

[English](README.md)

一套从零实现的浏览器 PPTX 内核，用来打开和编辑 `.pptx`。不使用 Univer，也不使用收费的 Office SDK。同一个页面可以单独打开，也可以作为 DeepSeek Harness（DSH）里 `.pptx` 的右侧预览。

## 预览

这个页面可以：

- 打开 `.pptx`，绘制背景、主题、母版、版式、文字、图片、表格、环形图和饼图
- 在幻灯片上直接改文字，并保留每一段文字原有的字体、字号、粗细和颜色
- 移动图片、文本框、环形图或饼图；调整图片、环形图和饼图的大小；裁剪图片
- 复制或删除选中的图片、文字、环形图和饼图；「裁剪」按钮的左上角贴在该图片的右上角
- 选中对象后用 Ctrl+C 复制、Ctrl+V 粘贴、Ctrl+Z 撤回上一步
- 用「插入图片」把本地 PNG、JPEG、GIF 或 WebP 放进当前页，再移动、缩放或裁剪
- 在末尾追加一张空白幻灯片，并沿用当前页的版式
- 复制幻灯片、删除幻灯片，或拖动缩略图调整顺序；最后一张不能删除
- 拖动缩略图和幻灯片之间的分隔条，同时改变两边的宽度
- 按 10% 到 400% 缩放，默认 100%；点「+」或「−」各变化 1%。只有选中幻灯片页面本身时，滚轮才会缩放
- 无论展示区多宽，横向滚动都能看到这一页的最左和最右
- 用「保存」覆盖当前打开的文件，用「放弃」回到最近一次保存的内容；还没保存过时，回到打开时的内容

除了环形图和饼图，其他图表、SmartArt、动画、切换、公式，以及和 Office 一致的换行，都还没有实现。文字换行是近似的。

## 本地运行

需要 Node.js `>=22.19.0` 和 `pnpm@11.23.0`。

```sh
pnpm install
pnpm run pptx:web
node packages/pptx-kernel/web/server.mjs
```

打开 `http://127.0.0.1:4173/`。「保存」会覆盖通过文件选择器、拖入的文件句柄，或本服务 `?open=` 打开的文件，不会另存下载。`POST /save?name=` 只接受 `packages/pptx-kernel/web/` 里已经存在的 `.pptx`。

`web/editor.js` 由 `pnpm run pptx:web` 生成，不要手改这个包。页面必须加载新包时，改 `packages/pptx-kernel/web/index.html` 里 `editor.js` 的查询参数。

## 安装到 DSH

在 DeepSeek Harness 源码目录把本仓库加进 Web 配置并启动：

```sh
pnpm dsh plugin --profile web add <本仓库路径>
pnpm dsh web
```

安装后重新打开 Web。DSH 里生成或点开的 `.pptx` 会用这个内核显示在会话右侧，和内置预览同一块区域。拖动这一栏的左边缘可以改变宽度。「打开方式」里仍然可以改回内置的 Office 预览。

登录后也可以打开 `http://127.0.0.1:3080/office/`。这个地址不接受 `/save`。在 DSH 预览里点「保存」，仍然需要先用「打开」选中文件，或使用独立预览服务提供的文件句柄。

## 开发

```sh
pnpm run typecheck
pnpm run test:unit
pnpm run lint
```

内核源码在 `packages/pptx-kernel/src`，页面在 `packages/pptx-kernel/web`。内核测试在 `test/unit/pptx-kernel.test.ts`。许可证是 Apache-2.0。
