import { readFileSync } from 'node:fs'
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { openPptx } from '../../packages/pptx-kernel/src/index.ts'
import { readZip, writeZip } from '../../packages/pptx-kernel/src/zip.ts'

const PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  )
)

describe('pptx kernel', () => {
  it('reads slides, edits one shape, and copies every other part through', async () => {
    const original = await sample()
    const opened = await openPptx(original)
    expect(opened.slideSize()).toEqual({ width: 12192000, height: 6858000 })
    const slides = opened.slides()
    expect(slides).toHaveLength(2)
    expect(slides[0]?.shapes.find((shape) => shape.id === '2')).toMatchObject({
      kind: 'text',
      name: 'Title',
      x: 100,
      y: 200,
      width: 300,
      height: 400,
      text: 'Hello & more'
    })
    expect(slides[0]?.shapes.find((shape) => shape.id === '3')).toMatchObject({
      kind: 'picture',
      imagePath: 'ppt/media/image1.png'
    })
    expect(opened.part('ppt/media/image1.png')).toEqual(PNG)
    expect(slides[0]?.shapes.find((shape) => shape.id === '5')).toMatchObject({
      text: '组内',
      x: 1050,
      y: 1060
    })
    expect(slides[1]?.shapes[0]?.text).toBe('第二页')
    const scene = opened.scene(0)
    expect(scene.background).toEqual({ kind: 'solid', color: '#ffffff' })
    const painted = scene.shapes.find((shape) => shape.id === '2')
    expect(painted?.plain).toBe('Hello & more')
    expect(painted?.text?.paragraphs[0]?.runs[0]).toMatchObject({
      fontSize: 18,
      color: '#000000',
      latin: 'Calibri'
    })
    expect(scene.shapes.find((shape) => shape.id === '5')).toMatchObject({ x: 1050, y: 1060 })
    expect(opened.scene(1).shapes[0]?.plain).toBe('第二页')

    const untouched = await opened.save()
    expect(untouched).toEqual(original)

    opened.setShapeText(0, '2', ' 改过了')
    expect(opened.scene(0).shapes.find((shape) => shape.id === '2')?.plain).toBe(' 改过了')
    const saved = await openPptx(await opened.save())
    const title = saved.slides()[0]?.shapes.find((shape) => shape.id === '2')
    expect(title?.text).toBe(' 改过了')
    expect(saved.part('ppt/media/image1.png')).toEqual(PNG)
    expect(
      new TextDecoder().decode(saved.part('ppt/notes/note.xml') ?? new Uint8Array())
    ).toContain('KEEP')
    expect(
      new TextDecoder().decode(saved.part('ppt/slides/slide1.xml') ?? new Uint8Array())
    ).toContain('xml:space="preserve"')
    expect(saved.slides()[1]?.shapes[0]?.text).toBe('第二页')
  })

  it('rejects a picture edit and a missing shape', async () => {
    const opened = await openPptx(await sample())
    expect(() => opened.setShapeText(0, '3', 'nope')).toThrow(/no text/)
    expect(() => opened.setShapeText(0, '99', 'nope')).toThrow(/does not exist/)
    expect(() => opened.setShapeText(4, '2', 'nope')).toThrow(/does not exist/)
  })

  it('paints an Office theme, master bar, and placeholder text', async () => {
    const opened = await openPptx(await themed())
    const scene = opened.scene(0)
    expect(scene.background).toEqual({ kind: 'solid', color: '#ff5500' })
    expect(scene.shapes[0]).toMatchObject({
      id: 'master:10',
      selectable: false,
      fill: { kind: 'solid', color: '#00aa00' },
      x: 0,
      y: 0,
      width: 9144000,
      height: 180000
    })
    const title = scene.shapes.find((shape) => shape.id === '2')
    expect(title).toMatchObject({
      plain: '季度汇报',
      x: 457200,
      y: 200000,
      width: 8000000,
      height: 1000000,
      editable: true
    })
    expect(title?.text?.anchor).toBe('ctr')
    expect(title?.text?.paragraphs[0]).toMatchObject({ align: 'ctr' })
    expect(title?.text?.paragraphs[0]?.runs[0]).toMatchObject({
      fontSize: 44,
      color: '#000000',
      latin: 'Calibri',
      eastAsian: '宋体'
    })
    expect(scene.shapes.find((shape) => shape.id === '3')).toMatchObject({
      geometry: 'round',
      corner: 0.5,
      fill: { kind: 'solid', color: '#a7c0de' },
      line: { color: '#112233', width: 12700 }
    })
    expect(scene.shapes.find((shape) => shape.id === '4')?.fill).toEqual({
      kind: 'linear',
      angle: 0,
      scaled: false,
      stops: [
        { offset: 0, color: '#ff0000' },
        { offset: 1, color: '#0000ff' }
      ]
    })
    const styled = scene.shapes.find((shape) => shape.id === '5')
    expect(styled?.fill).toEqual({ kind: 'solid', color: '#c0504d' })
    expect(styled?.text?.paragraphs[0]?.runs[0]).toMatchObject({
      fontSize: 18,
      color: '#ffffff',
      latin: 'Calibri'
    })
    expect(styled?.shadow?.color).toBe('rgba(0, 0, 0, 0.38)')
    expect(styled?.shadow?.offsetY).toBeCloseTo(20000, 0)
    expect(Math.abs(styled?.shadow?.offsetX ?? 1)).toBeLessThan(1)
    expect(scene.shapes.find((shape) => shape.id === '8')).toMatchObject({
      plain: 'Hi',
      x: 1200,
      y: 1000,
      width: 200,
      height: 50
    })
  })

  it('paints a PowerPoint slide with its theme background and picture', async () => {
    const opened = await openPptx(
      new Uint8Array(
        readFileSync(new URL('../fixtures/exchange/slide-picture.pptx', import.meta.url))
      )
    )
    const scene = opened.scene(0)
    expect(scene.background).toEqual({ kind: 'solid', color: '#ffffff' })
    expect(scene.shapes.some((shape) => shape.plain.includes('Click to edit'))).toBe(false)
    const picture = scene.shapes.find((shape) => shape.fill.kind === 'picture')
    expect(picture).toMatchObject({ x: 914400, y: 914400, width: 1828800, height: 1828800 })
    if (picture?.fill.kind === 'picture')
      expect(picture.fill.imagePath).toBe('ppt/media/image1.png')
    const date = scene.shapes.find((shape) => shape.plain.includes('1/27/13'))
    expect(date).toMatchObject({ x: 457200, y: 6356350, width: 2133600, height: 365125 })
    expect(date?.text?.paragraphs[0]?.runs[0]).toMatchObject({
      fontSize: 12,
      latin: 'Calibri',
      eastAsian: '宋体',
      color: '#404040'
    })
    expect(scene.shapes.find((shape) => shape.plain === '1')?.text?.paragraphs[0]?.align).toBe('r')
  })

  it('reads a deflated ZIP entry', async () => {
    const payload = new TextEncoder().encode('hello deflate')
    const archive = await writeZip([
      { name: 'ppt/presentation.xml', data: payload, method: 'deflate' }
    ])
    const parts = await readZip(archive)
    expect(parts.get('ppt/presentation.xml')).toEqual(payload)
    expect(deflateRawSync(payload).byteLength).toBeGreaterThan(0)
  })
})

async function sample(): Promise<Uint8Array> {
  const files: Array<[string, string | Uint8Array]> = [
    ['[Content_Types].xml', contentTypes()],
    ['_rels/.rels', rootRels()],
    ['ppt/presentation.xml', presentation()],
    ['ppt/_rels/presentation.xml.rels', presentationRels()],
    ['ppt/slides/slide1.xml', slideOne()],
    ['ppt/slides/_rels/slide1.xml.rels', slideRels()],
    ['ppt/slides/slide2.xml', slideTwo()],
    ['ppt/media/image1.png', PNG],
    ['ppt/notes/note.xml', '<note>KEEP</note>']
  ]
  return writeZip(
    files.map(([name, data]) => ({
      name,
      data: typeof data === 'string' ? new TextEncoder().encode(data) : data
    }))
  )
}

function contentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="png" ContentType="image/png"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`
}

function rootRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
}

function presentation(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldIdLst>
<p:sldId id="256" r:id="rId2"/>
<p:sldId id="257" r:id="rId3"/>
</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`
}

function presentationRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>`
}

function slideRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`
}

function slideOne(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>
<p:txBody><a:p><a:r><a:t>Hello</a:t></a:r><a:r><a:t> &amp; more</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:pic>
<p:nvPicPr><p:cNvPr id="3" name="Pic"/></p:nvPicPr>
<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
<p:spPr><a:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></a:xfrm></p:spPr>
</p:pic>
<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="4" name="Group"/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="2000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="2000"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="5" name="Inner"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="50" y="60"/><a:ext cx="10" cy="20"/></a:xfrm></p:spPr>
<p:txBody><a:p><a:r><a:t>组内</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:grpSp>
</p:spTree></p:cSld>
</p:sld>`
}

function slideTwo(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Body"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr>
<p:txBody><a:p><a:r><a:t>第二页</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:spTree></p:cSld>
</p:sld>`
}

async function themed(): Promise<Uint8Array> {
  const files: Array<[string, string]> = [
    ['ppt/presentation.xml', themedPresentation()],
    ['ppt/_rels/presentation.xml.rels', themedPresentationRels()],
    ['ppt/theme/theme1.xml', themedTheme()],
    ['ppt/slideMasters/slideMaster1.xml', themedMaster()],
    ['ppt/slideMasters/_rels/slideMaster1.xml.rels', themedMasterRels()],
    ['ppt/slideLayouts/slideLayout1.xml', themedLayout()],
    ['ppt/slideLayouts/_rels/slideLayout1.xml.rels', themedLayoutRels()],
    ['ppt/slides/slide1.xml', themedSlide()],
    ['ppt/slides/_rels/slide1.xml.rels', themedSlideRels()]
  ]
  return writeZip(files.map(([name, data]) => ({ name, data: new TextEncoder().encode(data) })))
}

function themedPresentation(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
<p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>`
}

function themedPresentationRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`
}

function themedTheme(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:srgbClr val="000000"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F497D"/></a:dk2>
<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
<a:accent2><a:srgbClr val="C0504D"/></a:accent2>
<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
<a:accent4><a:srgbClr val="8064A2"/></a:accent4>
<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
<a:accent6><a:srgbClr val="F79646"/></a:accent6>
<a:hlink><a:srgbClr val="0000FF"/></a:hlink>
<a:folHlink><a:srgbClr val="800080"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hans" typeface="宋体"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hans" typeface="宋体"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`
}

function themedMaster(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF5500"/></a:solidFill></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="10" name="Bar"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="180000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
</p:sp>
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:txStyles>
<p:titleStyle><a:lvl1pPr algn="ctr"><a:buNone/><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:buChar char="•"/><a:defRPr sz="3200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:bodyStyle>
<p:otherStyle><a:lvl1pPr algn="l"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:otherStyle>
</p:txStyles>
</p:sldMaster>`
}

function themedMasterRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`
}

function themedLayout(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="457200" y="200000"/><a:ext cx="8000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>
</p:sp>
</p:spTree></p:cSld>
</p:sldLayout>`
}

function themedLayoutRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`
}

function themedSlideRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
}

function themedSlide(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="标题 1"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>季度汇报</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="圆角"/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="400000" y="1600000"/><a:ext cx="2000000" cy="800000"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>
<a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill>
<a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>
</p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>按钮</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="4" name="渐变"/></p:nvSpPr>
<p:spPr>
<a:xfrm><a:off x="2600000" y="1600000"/><a:ext cx="2000000" cy="800000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>
<a:ln><a:noFill/></a:ln>
</p:spPr>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="5" name="样式"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="4800000" y="1600000"/><a:ext cx="2000000" cy="800000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:style>
<a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef>
<a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
</p:style>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>样式</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:grpSp>
<p:nvGrpSpPr><p:cNvPr id="7" name="组"/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="8" name="内"/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="100" y="0"/><a:ext cx="100" cy="50"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:grpSp>
</p:spTree></p:cSld>
</p:sld>`
}
