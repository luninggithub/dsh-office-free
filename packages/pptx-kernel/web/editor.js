// packages/pptx-kernel/src/xml.ts
function childElements(xml, from = 0, to = xml.length) {
  const result = [];
  let index = from;
  while (index < to) {
    const start = xml.indexOf("<", index);
    if (start < 0 || start >= to) break;
    const tag = readTag(xml, start);
    if (tag === void 0 || tag.kind === "close") {
      index = start + 1;
      continue;
    }
    if (tag.kind === "empty") {
      result.push({
        prefix: tag.prefix,
        name: tag.name,
        attrs: tag.attrs,
        start: tag.start,
        openEnd: tag.end,
        closeStart: tag.end,
        end: tag.end
      });
      index = tag.end;
      continue;
    }
    let depth = 1;
    let cursor = tag.end;
    let closeStart = tag.end;
    while (cursor < xml.length && depth > 0) {
      const next = xml.indexOf("<", cursor);
      if (next < 0) break;
      const inner = readTag(xml, next);
      if (inner === void 0) {
        cursor = next + 1;
        continue;
      }
      if (inner.prefix === tag.prefix && inner.name === tag.name) {
        if (inner.kind === "open") depth += 1;
        else if (inner.kind === "close") {
          depth -= 1;
          if (depth === 0) closeStart = inner.start;
        }
      }
      cursor = inner.end;
    }
    result.push({
      prefix: tag.prefix,
      name: tag.name,
      attrs: tag.attrs,
      start: tag.start,
      openEnd: tag.end,
      closeStart,
      end: cursor
    });
    index = cursor;
  }
  return result;
}
function findDescendant(xml, name, from = 0, to = xml.length) {
  for (const element of childElements(xml, from, to)) {
    if (element.name === name) return element;
    const nested = findDescendant(xml, name, element.openEnd, element.closeStart);
    if (nested !== void 0) return nested;
  }
  return void 0;
}
function parseAttributes(attrs) {
  const parsed = [];
  let index = 0;
  while (index < attrs.length) {
    while (index < attrs.length && /\s/.test(attrs[index] ?? "")) index += 1;
    if (index >= attrs.length) break;
    const nameStart = index;
    while (index < attrs.length && /[A-Za-z0-9_.:-]/.test(attrs[index] ?? "")) index += 1;
    const full = attrs.slice(nameStart, index);
    while (index < attrs.length && /\s/.test(attrs[index] ?? "")) index += 1;
    if (attrs[index] !== "=") continue;
    index += 1;
    while (index < attrs.length && /\s/.test(attrs[index] ?? "")) index += 1;
    const quote = attrs[index];
    if (quote !== '"' && quote !== "'") continue;
    const valueStart = index + 1;
    const valueEnd = attrs.indexOf(quote, valueStart);
    if (valueEnd < 0) break;
    const colon = full.indexOf(":");
    parsed.push({
      prefix: colon === -1 ? "" : full.slice(0, colon),
      name: colon === -1 ? full : full.slice(colon + 1),
      value: decodeXml(attrs.slice(valueStart, valueEnd))
    });
    index = valueEnd + 1;
  }
  return parsed;
}
function attribute(attrs, name, prefixed = false) {
  return parseAttributes(attrs).find(
    (item) => item.name === name && (prefixed ? item.prefix !== "" : item.prefix === "")
  )?.value;
}
function decodeXml(value) {
  return value.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (entity, token) => {
    if (token === "amp") return "&";
    if (token === "lt") return "<";
    if (token === "gt") return ">";
    if (token === "quot") return '"';
    if (token === "apos") return "'";
    const codePoint = token.startsWith("#x") ? Number.parseInt(token.slice(2), 16) : Number.parseInt(token.slice(1), 10);
    return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity;
  });
}
function escapeXml(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function readTag(xml, start) {
  if (xml[start] !== "<") return void 0;
  let index = start + 1;
  const marker = xml[index];
  if (marker === "?" || marker === "!") return void 0;
  let kind = "open";
  if (marker === "/") {
    kind = "close";
    index += 1;
  }
  const nameStart = index;
  while (index < xml.length && /[A-Za-z0-9_.:-]/.test(xml[index] ?? "")) index += 1;
  const full = xml.slice(nameStart, index);
  if (full.length === 0) return void 0;
  const colon = full.indexOf(":");
  const prefix = colon === -1 ? "" : full.slice(0, colon);
  const name = colon === -1 ? full : full.slice(colon + 1);
  const attrsStart = index;
  let quote = "";
  while (index < xml.length) {
    const char = xml[index] ?? "";
    if (quote !== "") {
      if (char === quote) quote = "";
      index += 1;
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      index += 1;
      continue;
    }
    if (char === ">") break;
    index += 1;
  }
  const empty = kind === "open" && xml[index - 1] === "/";
  const attrsEnd = empty ? index - 1 : index;
  return {
    kind: empty ? "empty" : kind,
    prefix,
    name,
    attrs: xml.slice(attrsStart, attrsEnd),
    start,
    end: index + 1
  };
}

// packages/pptx-kernel/src/geom.ts
var IDENTITY = { x: 0, y: 0, sx: 1, sy: 1 };
var EMPTY_XFRM = {
  x: 0,
  y: 0,
  width: 0,
  height: 0,
  chX: 0,
  chY: 0,
  chWidth: 0,
  chHeight: 0,
  rotation: 0,
  flipH: false,
  flipV: false
};
function readNumber(value, fallback = 0) {
  if (value === void 0) return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}
function shapeXfrm(xml, element) {
  const children = childElements(xml, element.openEnd, element.closeStart);
  const props = children.find((item) => item.name === "spPr" || item.name === "grpSpPr");
  if (props !== void 0) {
    const nested = findDescendant(xml, "xfrm", props.openEnd, props.closeStart);
    if (nested !== void 0) return xfrmFrom(xml, nested);
  }
  const own = children.find((item) => item.name === "xfrm");
  return own === void 0 ? void 0 : xfrmFrom(xml, own);
}
function mapBox(parent, x, y, width, height) {
  return {
    x: parent.x + x * parent.sx,
    y: parent.y + y * parent.sy,
    width: width * parent.sx,
    height: height * parent.sy
  };
}
function groupTransform(parent, group) {
  const box = mapBox(parent, group.x, group.y, group.width, group.height);
  const sx = group.chWidth === 0 ? parent.sx : box.width / group.chWidth;
  const sy = group.chHeight === 0 ? parent.sy : box.height / group.chHeight;
  return {
    x: box.x - group.chX * sx,
    y: box.y - group.chY * sy,
    sx,
    sy
  };
}
function rotatePoint(x, y, pivotX, pivotY, degrees) {
  if (degrees === 0) return { x, y };
  const rad = degrees * Math.PI / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = x - pivotX;
  const dy = y - pivotY;
  return {
    x: pivotX + dx * cos - dy * sin,
    y: pivotY + dx * sin + dy * cos
  };
}
function xfrmFrom(xml, element) {
  const kids = childElements(xml, element.openEnd, element.closeStart);
  const off = kids.find((item) => item.name === "off");
  const ext = kids.find((item) => item.name === "ext");
  const chOff = kids.find((item) => item.name === "chOff");
  const chExt = kids.find((item) => item.name === "chExt");
  return {
    x: readNumber(off === void 0 ? void 0 : attribute(off.attrs, "x")),
    y: readNumber(off === void 0 ? void 0 : attribute(off.attrs, "y")),
    width: readNumber(ext === void 0 ? void 0 : attribute(ext.attrs, "cx")),
    height: readNumber(ext === void 0 ? void 0 : attribute(ext.attrs, "cy")),
    chX: readNumber(chOff === void 0 ? void 0 : attribute(chOff.attrs, "x")),
    chY: readNumber(chOff === void 0 ? void 0 : attribute(chOff.attrs, "y")),
    chWidth: readNumber(chExt === void 0 ? void 0 : attribute(chExt.attrs, "cx")),
    chHeight: readNumber(chExt === void 0 ? void 0 : attribute(chExt.attrs, "cy")),
    rotation: readNumber(attribute(element.attrs, "rot")) / 6e4,
    flipH: attribute(element.attrs, "flipH") === "1",
    flipV: attribute(element.attrs, "flipV") === "1"
  };
}

// packages/pptx-kernel/src/rels.ts
function readRelations(xml) {
  const relations = /* @__PURE__ */ new Map();
  for (const element of childElements(xml)) {
    if (element.name !== "Relationship" && element.name !== "Relationships") continue;
    const nested = element.name === "Relationships" ? childElements(xml, element.openEnd, element.closeStart) : [element];
    for (const relation of nested) {
      if (relation.name !== "Relationship") continue;
      const id = attribute(relation.attrs, "Id");
      const type = attribute(relation.attrs, "Type");
      const target = attribute(relation.attrs, "Target");
      if (id !== void 0 && type !== void 0 && target !== void 0) {
        relations.set(id, { type, target });
      }
    }
  }
  return relations;
}
function resolvePart(base, target) {
  if (target.startsWith("/")) return target.slice(1);
  const stack = base.split("/").slice(0, -1);
  for (const piece of target.split("/")) {
    if (piece === "..") stack.pop();
    else if (piece !== "." && piece !== "") stack.push(piece);
  }
  return stack.join("/");
}

// packages/pptx-kernel/src/color.ts
var MAP_KEYS = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink"
];
var DEFAULT_HEX = {
  dk1: "000000",
  lt1: "FFFFFF",
  dk2: "1F497D",
  lt2: "EEECE1",
  accent1: "4F81BD",
  accent2: "C0504D",
  accent3: "9BBB59",
  accent4: "8064A2",
  accent5: "4BACC6",
  accent6: "F79646",
  hlink: "0000FF",
  folHlink: "800080"
};
var DEFAULT_FONT = { latin: "Calibri", ea: "", cs: "", hans: "\u5B8B\u4F53" };
var PRESET = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "00FF00",
  blue: "0000FF",
  yellow: "FFFF00",
  gray: "808080",
  grey: "808080"
};
var COLOR_TAGS = /* @__PURE__ */ new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr"]);
function defaultColorMap() {
  return {
    bg1: "lt1",
    tx1: "dk1",
    bg2: "lt2",
    tx2: "dk2",
    accent1: "accent1",
    accent2: "accent2",
    accent3: "accent3",
    accent4: "accent4",
    accent5: "accent5",
    accent6: "accent6",
    hlink: "hlink",
    folHlink: "folHlink"
  };
}
function colorMapFrom(attrs) {
  if (attrs === void 0) return void 0;
  const map = { ...defaultColorMap() };
  let seen = false;
  for (const key of MAP_KEYS) {
    const value = attribute(attrs, key);
    if (value === void 0) continue;
    map[key] = value;
    seen = true;
  }
  return seen ? map : void 0;
}
function readTheme(xml) {
  const colors = defaultPalette();
  if (xml === void 0) return emptyTheme(colors);
  const scheme = findDescendant(xml, "clrScheme");
  if (scheme !== void 0) {
    for (const entry of childElements(xml, scheme.openEnd, scheme.closeStart)) {
      const paint = readSchemeEntry(xml, entry);
      if (paint !== void 0) colors[entry.name] = paint;
    }
  }
  const fonts = findDescendant(xml, "fontScheme");
  const format = findDescendant(xml, "fmtScheme");
  return {
    xml,
    colors,
    major: readFonts(xml, fonts ? direct(xml, fonts, "majorFont") : void 0),
    minor: readFonts(xml, fonts ? direct(xml, fonts, "minorFont") : void 0),
    fillStyles: styleList(xml, format, "fillStyleLst"),
    lineStyles: styleList(xml, format, "lnStyleLst"),
    effectStyles: styleList(xml, format, "effectStyleLst"),
    bgStyles: styleList(xml, format, "bgFillStyleLst")
  };
}
function colorElement(xml, parent) {
  return childElements(xml, parent.openEnd, parent.closeStart).find(
    (item) => COLOR_TAGS.has(item.name)
  );
}
function schemeColor(value, ctx) {
  if (value === "phClr") return ctx.placeholder ?? { r: 0, g: 0, b: 0, a: 1 };
  const mapped = ctx.map[value] ?? value;
  return ctx.theme.colors[mapped] ?? ctx.theme.colors[value] ?? { r: 0, g: 0, b: 0, a: 1 };
}
function rgbOf(xml, element, ctx) {
  const value = attribute(element.attrs, "val") ?? "";
  let base;
  if (element.name === "srgbClr") base = hexRgb(value);
  else if (element.name === "schemeClr") base = schemeColor(value, ctx);
  else if (element.name === "sysClr") base = hexRgb(attribute(element.attrs, "lastClr") ?? "000000");
  else if (element.name === "prstClr") base = hexRgb(PRESET[value] ?? "000000");
  else if (element.name === "scrgbClr") base = scrgb(element.attrs);
  else base = { r: 0, g: 0, b: 0, a: 1 };
  return applyMods(base, childElements(xml, element.openEnd, element.closeStart));
}
function cssColor(color) {
  const r = clampByte(color.r);
  const g = clampByte(color.g);
  const b = clampByte(color.b);
  const alpha = Math.round(clamp01(color.a) * 1e3) / 1e3;
  if (alpha >= 0.999) return `#${hex(r)}${hex(g)}${hex(b)}`;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}
function direct(xml, element, name) {
  return childElements(xml, element.openEnd, element.closeStart).find((item) => item.name === name);
}
function emptyTheme(colors) {
  return {
    xml: "",
    colors,
    major: DEFAULT_FONT,
    minor: DEFAULT_FONT,
    fillStyles: [],
    lineStyles: [],
    effectStyles: [],
    bgStyles: []
  };
}
function defaultPalette() {
  const colors = {};
  for (const [name, value] of Object.entries(DEFAULT_HEX)) colors[name] = hexRgb(value);
  return colors;
}
function readSchemeEntry(xml, element) {
  const inner = colorElement(xml, element);
  if (inner === void 0) return void 0;
  if (inner.name === "sysClr") return hexRgb(attribute(inner.attrs, "lastClr") ?? "000000");
  if (inner.name === "srgbClr") return hexRgb(attribute(inner.attrs, "val") ?? "000000");
  return void 0;
}
function readFonts(xml, element) {
  if (element === void 0) return DEFAULT_FONT;
  const hans = childElements(xml, element.openEnd, element.closeStart).find(
    (item) => item.name === "font" && attribute(item.attrs, "script") === "Hans"
  );
  return {
    latin: face(xml, element, "latin") || DEFAULT_FONT.latin,
    ea: face(xml, element, "ea"),
    cs: face(xml, element, "cs"),
    hans: (hans === void 0 ? void 0 : attribute(hans.attrs, "typeface")) || DEFAULT_FONT.hans
  };
}
function face(xml, element, name) {
  const node = direct(xml, element, name);
  if (node === void 0) return "";
  return attribute(node.attrs, "typeface") ?? "";
}
function styleList(xml, format, name) {
  if (format === void 0) return [];
  const list = direct(xml, format, name);
  if (list === void 0) return [];
  return childElements(xml, list.openEnd, list.closeStart);
}
function scrgb(attrs) {
  const channel = (name) => readNumber(attribute(attrs, name)) / 1e5 * 255;
  return { r: channel("r"), g: channel("g"), b: channel("b"), a: 1 };
}
function applyMods(input, mods) {
  const color = { r: input.r, g: input.g, b: input.b, a: input.a };
  for (const mod of mods) {
    if (COLOR_TAGS.has(mod.name)) continue;
    const amount = readNumber(attribute(mod.attrs, "val")) / 1e5;
    if (mod.name === "alpha") color.a *= amount;
    else if (mod.name === "alphaOff") color.a = clamp01(color.a + amount);
    else if (mod.name === "shade") {
      color.r *= amount;
      color.g *= amount;
      color.b *= amount;
    } else if (mod.name === "tint") {
      color.r = color.r * amount + 255 * (1 - amount);
      color.g = color.g * amount + 255 * (1 - amount);
      color.b = color.b * amount + 255 * (1 - amount);
    } else if (mod.name === "lumMod" || mod.name === "lumOff" || mod.name === "satMod") {
      const next = hslAdjust(color, mod.name, amount);
      color.r = next.r;
      color.g = next.g;
      color.b = next.b;
      color.a = next.a;
    }
  }
  return color;
}
function hslAdjust(color, name, amount) {
  const hsl = rgbToHsl(color);
  let { s, l } = hsl;
  if (name === "lumMod") l *= amount;
  if (name === "lumOff") l += amount;
  if (name === "satMod") s *= amount;
  const next = hslToRgb(hsl.h, clamp01(s), clamp01(l));
  return { ...next, a: color.a };
}
function rgbToHsl(color) {
  const r = color.r / 255;
  const g = color.g / 255;
  const b = color.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const delta = max - min;
  const s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let h = 0;
  if (max === r) h = (g - b) / delta + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / delta + 2;
  else h = (r - g) / delta + 4;
  return { h: h * 60, s, l };
}
function hslToRgb(h, s, l) {
  if (s === 0) {
    const gray = l * 255;
    return { r: gray, g: gray, b: gray };
  }
  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const hp = h / 60;
  const x = chroma * (1 - Math.abs(hp % 2 - 1));
  const m = l - chroma / 2;
  let r = 0;
  let g = 0;
  let b = 0;
  if (hp < 1) {
    r = chroma;
    g = x;
  } else if (hp < 2) {
    r = x;
    g = chroma;
  } else if (hp < 3) {
    g = chroma;
    b = x;
  } else if (hp < 4) {
    g = x;
    b = chroma;
  } else if (hp < 5) {
    r = x;
    b = chroma;
  } else {
    r = chroma;
    b = x;
  }
  return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255 };
}
function hexRgb(text) {
  const value = text.trim();
  const r = Number.parseInt(value.slice(0, 2), 16);
  const g = Number.parseInt(value.slice(2, 4), 16);
  const b = Number.parseInt(value.slice(4, 6), 16);
  const a = value.length >= 8 ? Number.parseInt(value.slice(6, 8), 16) / 255 : 1;
  return {
    r: Number.isFinite(r) ? r : 0,
    g: Number.isFinite(g) ? g : 0,
    b: Number.isFinite(b) ? b : 0,
    a: Number.isFinite(a) ? a : 1
  };
}
function clampByte(value) {
  return Math.min(255, Math.max(0, Math.round(value)));
}
function clamp01(value) {
  return Math.min(1, Math.max(0, value));
}
function hex(value) {
  return clampByte(value).toString(16).padStart(2, "0");
}

// packages/pptx-kernel/src/scene.ts
var NO_FILL = { kind: "none" };
var NO_FONT = { latin: void 0, ea: void 0, color: void 0 };
function buildSlideScene(read, slideIndex) {
  const presentation = read("ppt/presentation.xml");
  if (presentation === void 0) throw new Error("pptx: presentation.xml is missing");
  const slides = slidePartNames(read, presentation);
  const partName = slides[slideIndex];
  if (partName === void 0) throw new Error(`pptx: slide ${slideIndex} does not exist`);
  const slideXml = read(partName);
  if (slideXml === void 0) throw new Error(`pptx: part ${partName} is missing`);
  const slideRels = partRelations(read, partName);
  const layoutName = targetByType(slideRels, partName, "/slideLayout");
  const layoutXml = layoutName === void 0 ? void 0 : read(layoutName);
  const layoutRels = layoutName === void 0 ? /* @__PURE__ */ new Map() : partRelations(read, layoutName);
  const masterName = layoutName === void 0 ? void 0 : targetByType(layoutRels, layoutName, "/slideMaster");
  const masterXml = masterName === void 0 ? void 0 : read(masterName);
  const masterRels = masterName === void 0 ? /* @__PURE__ */ new Map() : partRelations(read, masterName);
  const themeName = masterName === void 0 ? void 0 : targetByType(masterRels, masterName, "/theme");
  const theme = readTheme(themeName === void 0 ? void 0 : read(themeName));
  const map = effectiveMap(slideXml, layoutXml, masterXml);
  const styles = textStyles(presentation, masterXml);
  const layoutMarkers = layoutXml === void 0 ? [] : markersOf(layoutXml);
  const masterMarkers = masterXml === void 0 ? [] : markersOf(masterXml);
  const slideMarkers = markersOf(slideXml);
  const images = [];
  const deck = {
    theme,
    map,
    ctx: { theme, map, placeholder: void 0 },
    images,
    fallback: styles.fallback,
    title: styles.title,
    body: styles.body,
    other: styles.other,
    layoutXml,
    masterXml,
    layoutMarkers,
    masterMarkers,
    slideKinds: kindsOf(slideMarkers),
    layoutKinds: kindsOf(layoutMarkers),
    hf: headerFooter(slideXml, layoutXml, masterXml),
    slideNumber: slideIndex + 1
  };
  const size = readSlideSize(presentation);
  const shapes = [];
  const showLayout = showsMaster(slideXml);
  const showMasterShapes = showLayout && (layoutXml === void 0 || showsMaster(layoutXml));
  if (showMasterShapes && masterXml !== void 0 && masterName !== void 0) {
    paintPart(deck, masterXml, masterName, masterRels, "master", shapes);
  }
  if (showLayout && layoutXml !== void 0 && layoutName !== void 0) {
    paintPart(deck, layoutXml, layoutName, layoutRels, "layout", shapes);
  }
  paintPart(deck, slideXml, partName, slideRels, "slide", shapes);
  return {
    width: size.width,
    height: size.height,
    background: resolveBackground(
      deck,
      slideXml,
      partName,
      slideRels,
      layoutXml,
      layoutName,
      layoutRels,
      masterXml,
      masterName,
      masterRels
    ),
    shapes,
    imagePaths: images
  };
}
function slidePartNames(read, presentation) {
  const relations = readRelations(read("ppt/_rels/presentation.xml.rels") ?? "");
  const list = findDescendant(presentation, "sldIdLst");
  if (list === void 0) return [];
  const names = [];
  for (const slideId of childElements(presentation, list.openEnd, list.closeStart)) {
    if (slideId.name !== "sldId") continue;
    const relationId = attribute(slideId.attrs, "id", true);
    const target = relationId === void 0 ? void 0 : relations.get(relationId);
    if (target === void 0 || !target.type.endsWith("/slide")) continue;
    names.push(resolvePart("ppt/presentation.xml", target.target));
  }
  return names;
}
function partRelations(read, partName) {
  const slash = partName.lastIndexOf("/");
  const rels = `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
  const xml = read(rels);
  return xml === void 0 ? /* @__PURE__ */ new Map() : readRelations(xml);
}
function targetByType(rels, partName, suffix) {
  for (const relation of rels.values()) {
    if (relation.type.endsWith(suffix)) return resolvePart(partName, relation.target);
  }
  return void 0;
}
function readSlideSize(xml) {
  const element = findDescendant(xml, "sldSz");
  const width = readNumber(element === void 0 ? void 0 : attribute(element.attrs, "cx"));
  const height = readNumber(element === void 0 ? void 0 : attribute(element.attrs, "cy"));
  return {
    width: width > 0 ? width : 12192e3,
    height: height > 0 ? height : 6858e3
  };
}
function effectiveMap(slideXml, layoutXml, masterXml) {
  const master = masterXml === void 0 ? void 0 : findDescendant(masterXml, "clrMap");
  let map = colorMapFrom(master?.attrs) ?? defaultColorMap();
  const layout = layoutXml === void 0 ? void 0 : overrideMap(layoutXml);
  if (layout !== void 0) map = layout;
  const slide = overrideMap(slideXml);
  if (slide !== void 0) map = slide;
  return map;
}
function overrideMap(xml) {
  const host = findDescendant(xml, "clrMapOvr");
  if (host === void 0) return void 0;
  const over = direct2(xml, host, "overrideClrMapping");
  if (over === void 0) return void 0;
  return colorMapFrom(over.attrs);
}
function textStyles(presentation, masterXml) {
  const fallback = findDescendant(presentation, "defaultTextStyle");
  const tx = masterXml === void 0 ? void 0 : findDescendant(masterXml, "txStyles");
  return {
    fallback: fallback === void 0 ? void 0 : { xml: presentation, element: fallback },
    title: tx === void 0 || masterXml === void 0 ? void 0 : spanChild(masterXml, tx, "titleStyle"),
    body: tx === void 0 || masterXml === void 0 ? void 0 : spanChild(masterXml, tx, "bodyStyle"),
    other: tx === void 0 || masterXml === void 0 ? void 0 : spanChild(masterXml, tx, "otherStyle")
  };
}
function spanChild(xml, parent, name) {
  const element = direct2(xml, parent, name);
  return element === void 0 ? void 0 : { xml, element };
}
function headerFooter(slideXml, layoutXml, masterXml) {
  const element = findDescendant(slideXml, "hf") ?? (layoutXml === void 0 ? void 0 : findDescendant(layoutXml, "hf")) ?? (masterXml === void 0 ? void 0 : findDescendant(masterXml, "hf"));
  if (element === void 0) return { dt: true, ftr: true, sldNum: true };
  return {
    dt: attribute(element.attrs, "dt") !== "0",
    ftr: attribute(element.attrs, "ftr") !== "0",
    sldNum: attribute(element.attrs, "sldNum") !== "0"
  };
}
function showsMaster(xml) {
  const cSld = findDescendant(xml, "cSld");
  if (cSld === void 0) return true;
  return attribute(cSld.attrs, "showMasterSp") !== "0";
}
function resolveBackground(deck, slideXml, slideName, slideRels, layoutXml, layoutName, layoutRels, masterXml, masterName, masterRels) {
  const own = readBackground(slideXml, slideName, slideRels, deck);
  if (own !== void 0) return own;
  if (layoutXml !== void 0 && layoutName !== void 0) {
    const layout = readBackground(layoutXml, layoutName, layoutRels, deck);
    if (layout !== void 0) return layout;
  }
  if (masterXml !== void 0 && masterName !== void 0) {
    const master = readBackground(masterXml, masterName, masterRels, deck);
    if (master !== void 0) return master;
  }
  return { kind: "solid", color: "#ffffff" };
}
function readBackground(xml, partName, rels, deck) {
  const cSld = findDescendant(xml, "cSld");
  if (cSld === void 0) return void 0;
  const bg = direct2(xml, cSld, "bg");
  if (bg === void 0) return void 0;
  const ref = direct2(xml, bg, "bgRef");
  if (ref !== void 0) return fillFromRef(xml, ref, deck, partName, rels, "bg");
  const props = direct2(xml, bg, "bgPr");
  if (props === void 0) return void 0;
  return containerFill(xml, props, deck.ctx, deck, partName, rels, NO_FILL) ?? NO_FILL;
}
function paintPart(deck, xml, partName, rels, mode, shapes) {
  const tree = findDescendant(xml, "spTree");
  if (tree === void 0) return;
  paintTree(deck, xml, tree.openEnd, tree.closeStart, shapes, {
    frame: IDENTITY,
    rotation: 0,
    pivotX: 0,
    pivotY: 0,
    groupFill: NO_FILL,
    mode,
    partName,
    rels
  });
}
function paintTree(deck, xml, from, to, shapes, walk) {
  for (const element of childElements(xml, from, to)) {
    if (element.name === "grpSp") {
      if (isHidden(xml, element)) continue;
      const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM;
      const box = mapBox(walk.frame, xfrm.x, xfrm.y, xfrm.width, xfrm.height);
      const center = rotatePoint(
        box.x + box.width / 2,
        box.y + box.height / 2,
        walk.pivotX,
        walk.pivotY,
        walk.rotation
      );
      const props = direct2(xml, element, "grpSpPr");
      const own = props === void 0 ? void 0 : containerFill(xml, props, deck.ctx, deck, walk.partName, walk.rels, walk.groupFill);
      paintTree(deck, xml, element.openEnd, element.closeStart, shapes, {
        frame: groupTransform(walk.frame, xfrm),
        rotation: walk.rotation + xfrm.rotation,
        pivotX: walk.rotation === 0 ? center.x : walk.pivotX,
        pivotY: walk.rotation === 0 ? center.y : walk.pivotY,
        groupFill: own === void 0 || own.kind === "none" ? walk.groupFill : own,
        mode: walk.mode,
        partName: walk.partName,
        rels: walk.rels
      });
      continue;
    }
    if (element.name !== "sp" && element.name !== "pic" && element.name !== "cxnSp" && element.name !== "graphicFrame") {
      continue;
    }
    if (isHidden(xml, element)) continue;
    const marker = markerOf(xml, element);
    if (!includeShape(marker, walk.mode, deck)) continue;
    if (element.name === "graphicFrame") {
      paintTable(deck, xml, element, shapes, walk);
      continue;
    }
    const shape = buildShape(deck, xml, element, marker, walk);
    if (shape !== void 0) shapes.push(shape);
  }
}
function includeShape(marker, mode, deck) {
  if (mode === "slide" || marker === void 0) return true;
  const kind = alias(marker.type);
  if (kind !== "dt" && kind !== "ftr" && kind !== "sldNum") return false;
  if (kind === "dt" && !deck.hf.dt) return false;
  if (kind === "ftr" && !deck.hf.ftr) return false;
  if (kind === "sldNum" && !deck.hf.sldNum) return false;
  if (mode === "layout" && deck.slideKinds.has(kind)) return false;
  if (mode === "master" && (deck.slideKinds.has(kind) || deck.layoutKinds.has(kind))) return false;
  return true;
}
function buildShape(deck, xml, element, marker, walk) {
  const identity = identityOf(xml, element);
  if (identity === void 0) return void 0;
  const local = inheritedXfrm(deck, xml, element, marker, walk.mode);
  if (local === void 0) return void 0;
  const placed = place(walk, local.x, local.y, local.width, local.height, local.rotation);
  if (placed.width === 0 && placed.height === 0) return void 0;
  const visual = visuals(deck, xml, element, walk, walk.groupFill);
  const text = element.name === "pic" ? void 0 : paintText(deck, xml, element, marker, walk.mode);
  return {
    id: walk.mode === "slide" ? identity.id : `${walk.mode}:${identity.id}`,
    name: identity.name,
    x: placed.x,
    y: placed.y,
    width: placed.width,
    height: placed.height,
    rotation: placed.rotation,
    flipH: local.flipH,
    flipV: local.flipV,
    geometry: visual.geometry,
    corner: visual.corner,
    fill: visual.fill,
    line: visual.line,
    shadow: visual.shadow,
    text,
    plain: plainOf(text),
    editable: walk.mode === "slide" && text !== void 0,
    selectable: walk.mode === "slide"
  };
}
function place(walk, x, y, width, height, rotation) {
  const placed = mapBox(walk.frame, x, y, width, height);
  const center = rotatePoint(
    placed.x + placed.width / 2,
    placed.y + placed.height / 2,
    walk.pivotX,
    walk.pivotY,
    walk.rotation
  );
  return {
    x: center.x - placed.width / 2,
    y: center.y - placed.height / 2,
    width: placed.width,
    height: placed.height,
    rotation: walk.rotation + rotation
  };
}
function inheritedXfrm(deck, xml, element, marker, mode) {
  const own = shapeXfrm(xml, element);
  if (own !== void 0) return own;
  const proto = prototypes(deck, marker, mode);
  if (proto.layout !== void 0 && deck.layoutXml !== void 0) {
    const layout = shapeXfrm(deck.layoutXml, proto.layout.element);
    if (layout !== void 0) return layout;
  }
  if (proto.master !== void 0 && deck.masterXml !== void 0) {
    return shapeXfrm(deck.masterXml, proto.master.element);
  }
  return void 0;
}
function visuals(deck, xml, element, walk, groupFill) {
  const spPr = direct2(xml, element, "spPr");
  const style = direct2(xml, element, "style");
  const geometry = geometryOf(deck, xml, element, markerOf(xml, element), walk.mode);
  let fill;
  if (element.name === "pic") {
    const blip = direct2(xml, element, "blipFill");
    fill = blip === void 0 ? NO_FILL : paintFillElement(xml, blip, deck.ctx, deck, walk.partName, walk.rels, groupFill);
  } else if (spPr !== void 0) {
    fill = containerFill(xml, spPr, deck.ctx, deck, walk.partName, walk.rels, groupFill);
  }
  if (fill === void 0 && style !== void 0) {
    const ref = direct2(xml, style, "fillRef");
    if (ref !== void 0) fill = fillFromRef(xml, ref, deck, walk.partName, walk.rels, "fill");
  }
  return {
    geometry: geometry.geometry,
    corner: geometry.corner,
    fill: fill ?? NO_FILL,
    line: shapeLine(xml, spPr, style, deck),
    shadow: shapeShadow(xml, spPr, style, deck)
  };
}
function geometryOf(deck, xml, element, marker, mode) {
  const proto = prototypes(deck, marker, mode);
  const spans = [];
  if (proto.master !== void 0 && deck.masterXml !== void 0) {
    spans.push({ xml: deck.masterXml, element: proto.master.element });
  }
  if (proto.layout !== void 0 && deck.layoutXml !== void 0) {
    spans.push({ xml: deck.layoutXml, element: proto.layout.element });
  }
  spans.push({ xml, element });
  let geometry = "rect";
  let corner = 0;
  for (const span of spans) {
    const spPr = direct2(span.xml, span.element, "spPr");
    const geom = spPr === void 0 ? void 0 : direct2(span.xml, spPr, "prstGeom");
    if (geom === void 0) continue;
    geometry = geometryKind(attribute(geom.attrs, "prst") ?? "rect");
    corner = geometry === "round" ? cornerOf(span.xml, geom) : 0;
  }
  return { geometry, corner };
}
function geometryKind(prst) {
  if (prst === "ellipse" || prst === "oval") return "ellipse";
  if (prst === "triangle" || prst === "isoscelesTriangle") return "triangle";
  if (prst === "diamond") return "diamond";
  if (prst === "line" || prst === "straightConnector1") return "line";
  if (prst.includes("round") && prst.includes("Rect")) return "round";
  return "rect";
}
function cornerOf(xml, geom) {
  const list = direct2(xml, geom, "avLst");
  const guides = list === void 0 ? [] : childElements(xml, list.openEnd, list.closeStart);
  const adj = guides.find((item) => attribute(item.attrs, "name") === "adj");
  const formula = adj === void 0 ? "" : attribute(adj.attrs, "fmla") ?? "";
  const match = /^val\s+(-?\d+(?:\.\d+)?)/.exec(formula);
  if (match?.[1] !== void 0) return readNumber(match[1]) / 1e5;
  return 16667 / 1e5;
}
function paintTable(deck, xml, element, shapes, walk) {
  const table = findDescendant(xml, "tbl", element.openEnd, element.closeStart);
  if (table === void 0) return;
  const identity = identityOf(xml, element);
  const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM;
  const grid = direct2(xml, table, "tblGrid");
  const cols = grid === void 0 ? [] : childElements(xml, grid.openEnd, grid.closeStart).filter((item) => item.name === "gridCol").map((item) => readNumber(attribute(item.attrs, "w")));
  const rows = childElements(xml, table.openEnd, table.closeStart).filter(
    (item) => item.name === "tr"
  );
  let top = xfrm.y;
  rows.forEach((row, rowIndex) => {
    const height = readNumber(attribute(row.attrs, "h"));
    let left = xfrm.x;
    let col = 0;
    const cells = childElements(xml, row.openEnd, row.closeStart).filter(
      (item) => item.name === "tc"
    );
    for (const cell of cells) {
      if (attribute(cell.attrs, "hMerge") === "1") continue;
      const span = Math.max(1, readNumber(attribute(cell.attrs, "gridSpan"), 1));
      const width = sumCols(cols, col, span);
      if (attribute(cell.attrs, "vMerge") !== "1" && width > 0 && height > 0) {
        const placed = place(walk, left, top, width, height, xfrm.rotation);
        const text = cellText(deck, xml, cell);
        const props = direct2(xml, cell, "tcPr");
        const fill = props === void 0 ? NO_FILL : containerFill(xml, props, deck.ctx, deck, walk.partName, walk.rels, NO_FILL) ?? NO_FILL;
        shapes.push({
          id: `${identity?.id ?? "table"}:${rowIndex}:${col}`,
          name: identity?.name ?? "",
          x: placed.x,
          y: placed.y,
          width: placed.width,
          height: placed.height,
          rotation: placed.rotation,
          flipH: false,
          flipV: false,
          geometry: "rect",
          corner: 0,
          fill,
          line: props === void 0 ? void 0 : cellLine(xml, props, deck),
          shadow: void 0,
          text,
          plain: plainOf(text),
          editable: false,
          selectable: false
        });
      }
      left += width;
      col += span;
    }
    top += height;
  });
}
function sumCols(cols, start, span) {
  let total = 0;
  for (let index = 0; index < span; index += 1) total += cols[start + index] ?? 0;
  return total;
}
function cellLine(xml, props, deck) {
  const edge = ["lnL", "lnR", "lnT", "lnB"].map((name) => direct2(xml, props, name)).find((item) => item !== void 0);
  if (edge === void 0) return void 0;
  return lineFrom(xml, edge, deck.ctx);
}
function cellText(deck, xml, cell) {
  const body = direct2(xml, cell, "txBody");
  if (body === void 0) return void 0;
  const own = direct2(xml, body, "lstStyle");
  const rest = own === void 0 ? [] : [{ xml, element: own }];
  const props = direct2(xml, cell, "tcPr");
  return paragraphs(deck, xml, body, deck.fallback, NO_FONT, rest, paddingOf(props), 1);
}
function paddingOf(props) {
  return {
    left: readNumber(props === void 0 ? void 0 : attribute(props.attrs, "marL"), 91440),
    top: readNumber(props === void 0 ? void 0 : attribute(props.attrs, "marT"), 45720),
    right: readNumber(props === void 0 ? void 0 : attribute(props.attrs, "marR"), 91440),
    bottom: readNumber(props === void 0 ? void 0 : attribute(props.attrs, "marB"), 45720)
  };
}
function paintText(deck, xml, element, marker, mode) {
  const body = direct2(xml, element, "txBody");
  if (body === void 0) return void 0;
  const metrics = bodyMetrics(deck, xml, element, marker, mode);
  return paragraphs(
    deck,
    xml,
    body,
    deck.fallback,
    fontRefOf(xml, element, deck.ctx),
    styleSpans(deck, xml, element, marker, mode),
    metrics.padding,
    metrics.fontScale,
    metrics.anchor
  );
}
function paragraphs(deck, xml, body, fallback, font, rest, padding, fontScale, anchor = "t") {
  const built = [];
  const counters = /* @__PURE__ */ new Map();
  for (const paragraph of childElements(xml, body.openEnd, body.closeStart)) {
    if (paragraph.name !== "p") continue;
    const props = direct2(xml, paragraph, "pPr");
    const level = props === void 0 ? 0 : readNumber(attribute(props.attrs, "lvl"));
    const drafted = draftFor(deck, level, fallback, font, rest);
    if (props !== void 0) applyParaProps(xml, props, drafted.run, drafted.para, deck.ctx);
    const runs = runsOf(xml, paragraph, drafted.run, deck, fontScale);
    if (runs.length === 0) {
      const end = direct2(xml, paragraph, "endParaRPr");
      const next = { ...drafted.run };
      if (end !== void 0) applyRun(xml, end, next, deck.ctx);
      runs.push(toRun(next, "", deck.theme, fontScale));
    }
    const first = runs[0];
    const fontPt = first === void 0 ? drafted.run.sizePt * fontScale : first.fontSize;
    built.push({
      align: drafted.para.align,
      marginLeft: drafted.para.marginLeft,
      indent: drafted.para.indent,
      spaceBefore: spacePt(drafted.para.before, fontPt),
      spaceAfter: spacePt(drafted.para.after, fontPt),
      linePct: drafted.para.linePct,
      linePt: drafted.para.linePt,
      bullet: bulletOf(drafted.para.bullet, level, counters),
      runs
    });
  }
  return { anchor, padding, paragraphs: built };
}
function styleSpans(deck, xml, element, marker, mode) {
  const spans = [];
  const bucket = bucketSpan(deck, marker?.type);
  if (bucket !== void 0) spans.push(bucket);
  const proto = prototypes(deck, marker, mode);
  if (proto.master !== void 0 && deck.masterXml !== void 0) {
    const list = listStyle(deck.masterXml, proto.master.element);
    if (list !== void 0) spans.push({ xml: deck.masterXml, element: list });
  }
  if (proto.layout !== void 0 && deck.layoutXml !== void 0) {
    const list = listStyle(deck.layoutXml, proto.layout.element);
    if (list !== void 0) spans.push({ xml: deck.layoutXml, element: list });
  }
  const own = listStyle(xml, element);
  if (own !== void 0) spans.push({ xml, element: own });
  return spans;
}
function bucketSpan(deck, type) {
  if (type === void 0) return void 0;
  const kind = alias(type);
  if (kind === "title") return deck.title;
  if (kind === "body") return deck.body;
  return deck.other;
}
function listStyle(xml, element) {
  const body = direct2(xml, element, "txBody");
  if (body === void 0) return void 0;
  return direct2(xml, body, "lstStyle");
}
function bodyMetrics(deck, xml, element, marker, mode) {
  let anchor = "t";
  let left = 91440;
  let top = 45720;
  let right = 91440;
  let bottom = 45720;
  let fontScale = 1;
  const proto = prototypes(deck, marker, mode);
  const spans = [];
  if (proto.master !== void 0 && deck.masterXml !== void 0) {
    spans.push({ xml: deck.masterXml, element: proto.master.element });
  }
  if (proto.layout !== void 0 && deck.layoutXml !== void 0) {
    spans.push({ xml: deck.layoutXml, element: proto.layout.element });
  }
  spans.push({ xml, element });
  for (const span of spans) {
    const body = direct2(span.xml, span.element, "txBody");
    const props = body === void 0 ? void 0 : direct2(span.xml, body, "bodyPr");
    if (props === void 0) continue;
    const value = attribute(props.attrs, "anchor");
    if (value === "ctr" || value === "b" || value === "t") anchor = value;
    const lIns = attribute(props.attrs, "lIns");
    if (lIns !== void 0) left = readNumber(lIns);
    const tIns = attribute(props.attrs, "tIns");
    if (tIns !== void 0) top = readNumber(tIns);
    const rIns = attribute(props.attrs, "rIns");
    if (rIns !== void 0) right = readNumber(rIns);
    const bIns = attribute(props.attrs, "bIns");
    if (bIns !== void 0) bottom = readNumber(bIns);
    const fit = direct2(span.xml, props, "normAutofit");
    if (fit !== void 0) {
      const scale = attribute(fit.attrs, "fontScale");
      fontScale = scale === void 0 ? 1 : readNumber(scale) / 1e5;
    }
  }
  return { anchor, padding: { left, top, right, bottom }, fontScale };
}
function prototypes(deck, marker, mode) {
  if (marker === void 0 || mode === "master") return { layout: void 0, master: void 0 };
  if (mode === "layout") {
    return {
      layout: void 0,
      master: deck.masterXml === void 0 ? void 0 : findPrototype(deck.masterMarkers, marker)
    };
  }
  const layout = deck.layoutXml === void 0 ? void 0 : findPrototype(deck.layoutMarkers, marker);
  const masterKey = layout ?? marker;
  return {
    layout,
    master: deck.masterXml === void 0 ? void 0 : findPrototype(deck.masterMarkers, masterKey)
  };
}
function draftFor(deck, level, fallback, font, rest) {
  const run = {
    sizePt: 18,
    bold: false,
    italic: false,
    underline: false,
    color: cssColor(schemeColor("tx1", deck.ctx)),
    latin: "+mn-lt",
    ea: "+mn-ea"
  };
  const para = emptyPara();
  if (fallback !== void 0) applyLevel(fallback.xml, fallback.element, level, run, para, deck.ctx);
  if (font.latin !== void 0) run.latin = font.latin;
  if (font.ea !== void 0) run.ea = font.ea;
  if (font.color !== void 0) run.color = font.color;
  for (const span of rest) applyLevel(span.xml, span.element, level, run, para, deck.ctx);
  return { run, para };
}
function emptyPara() {
  return {
    align: "l",
    marginLeft: 0,
    indent: 0,
    bullet: void 0,
    before: { kind: "pt", value: 0 },
    after: { kind: "pt", value: 0 },
    linePct: 1,
    linePt: 0
  };
}
function applyLevel(xml, container, level, run, para, ctx) {
  const element = levelElement(xml, container, level);
  if (element === void 0) return;
  applyParaProps(xml, element, run, para, ctx);
}
function levelElement(xml, container, level) {
  const children = childElements(xml, container.openEnd, container.closeStart);
  for (let index = level; index >= 0; index -= 1) {
    const found = children.find((item) => item.name === `lvl${index + 1}pPr`);
    if (found !== void 0) return found;
  }
  return children.find((item) => item.name === "defPPr");
}
function applyParaProps(xml, element, run, para, ctx) {
  const align = attribute(element.attrs, "algn");
  if (align === "l" || align === "ctr" || align === "r") para.align = align;
  if (align === "just") para.align = "l";
  const margin = attribute(element.attrs, "marL");
  if (margin !== void 0) para.marginLeft = readNumber(margin);
  const indent = attribute(element.attrs, "indent");
  if (indent !== void 0) para.indent = readNumber(indent);
  const kids = childElements(xml, element.openEnd, element.closeStart);
  const bullet = kids.find(
    (item) => item.name === "buNone" || item.name === "buChar" || item.name === "buAutoNum"
  );
  if (bullet?.name === "buNone") para.bullet = void 0;
  if (bullet?.name === "buChar") para.bullet = attribute(bullet.attrs, "char") ?? "\u2022";
  if (bullet?.name === "buAutoNum") para.bullet = "auto";
  const before = kids.find((item) => item.name === "spcBef");
  if (before !== void 0) para.before = readSpace(xml, before);
  const after = kids.find((item) => item.name === "spcAft");
  if (after !== void 0) para.after = readSpace(xml, after);
  const line = kids.find((item) => item.name === "lnSpc");
  if (line !== void 0) {
    const spacing = readLine(xml, line);
    para.linePct = spacing.linePct;
    para.linePt = spacing.linePt;
  }
  const props = kids.find((item) => item.name === "defRPr");
  if (props !== void 0) applyRun(xml, props, run, ctx);
}
function readSpace(xml, element) {
  const pct = direct2(xml, element, "spcPct");
  if (pct !== void 0)
    return { kind: "pct", value: readNumber(attribute(pct.attrs, "val")) / 1e5 };
  const points = direct2(xml, element, "spcPts");
  if (points !== void 0) {
    return { kind: "pt", value: readNumber(attribute(points.attrs, "val")) / 100 };
  }
  return { kind: "pt", value: 0 };
}
function readLine(xml, element) {
  const pct = direct2(xml, element, "spcPct");
  if (pct !== void 0) {
    return { linePct: readNumber(attribute(pct.attrs, "val")) / 1e5, linePt: 0 };
  }
  const points = direct2(xml, element, "spcPts");
  if (points !== void 0) {
    return { linePct: 1, linePt: readNumber(attribute(points.attrs, "val")) / 100 };
  }
  return { linePct: 1, linePt: 0 };
}
function spacePt(space, fontPt) {
  return space.kind === "pt" ? space.value : space.value * fontPt;
}
function runsOf(xml, paragraph, base, deck, fontScale) {
  const runs = [];
  for (const child of childElements(xml, paragraph.openEnd, paragraph.closeStart)) {
    if (child.name !== "r" && child.name !== "fld" && child.name !== "br") continue;
    const next = { ...base };
    const props = direct2(xml, child, "rPr");
    if (props !== void 0) applyRun(xml, props, next, deck.ctx);
    if (child.name === "br") {
      runs.push(toRun(next, "\n", deck.theme, fontScale));
      continue;
    }
    const text = child.name === "fld" && attribute(child.attrs, "type") === "slidenum" ? String(deck.slideNumber) : inlineText(xml, child);
    runs.push(toRun(next, text, deck.theme, fontScale));
  }
  return runs;
}
function applyRun(xml, props, run, ctx) {
  const size = attribute(props.attrs, "sz");
  if (size !== void 0) run.sizePt = readNumber(size) / 100;
  const bold = attribute(props.attrs, "b");
  if (bold === "1" || bold === "true") run.bold = true;
  if (bold === "0" || bold === "false") run.bold = false;
  const italic = attribute(props.attrs, "i");
  if (italic === "1" || italic === "true") run.italic = true;
  if (italic === "0" || italic === "false") run.italic = false;
  const underline = attribute(props.attrs, "u");
  if (underline !== void 0) run.underline = underline !== "none";
  const kids = childElements(xml, props.openEnd, props.closeStart);
  const solid = kids.find((item) => item.name === "solidFill");
  if (solid !== void 0) {
    const color = colorElement(xml, solid);
    if (color !== void 0) run.color = cssColor(rgbOf(xml, color, ctx));
  }
  const latin = kids.find((item) => item.name === "latin");
  const latinFace = latin === void 0 ? void 0 : attribute(latin.attrs, "typeface");
  if (latinFace !== void 0 && latinFace !== "") run.latin = latinFace;
  const ea = kids.find((item) => item.name === "ea");
  const eaFace = ea === void 0 ? void 0 : attribute(ea.attrs, "typeface");
  if (eaFace !== void 0 && eaFace !== "") run.ea = eaFace;
}
function toRun(run, text, theme, fontScale) {
  return {
    text,
    fontSize: run.sizePt * fontScale,
    bold: run.bold,
    italic: run.italic,
    underline: run.underline,
    color: run.color,
    latin: resolveFace(run.latin, theme),
    eastAsian: resolveFace(run.ea, theme)
  };
}
function resolveFace(token, theme) {
  if (token === "+mj-lt") return theme.major.latin || "Calibri";
  if (token === "+mn-lt") return theme.minor.latin || "Calibri";
  if (token === "+mj-ea") return theme.major.ea || theme.major.hans || theme.major.latin || "\u5B8B\u4F53";
  if (token === "+mn-ea") return theme.minor.ea || theme.minor.hans || theme.minor.latin || "\u5B8B\u4F53";
  if (token === "+mj-cs") return theme.major.cs || theme.major.latin || "Calibri";
  if (token === "+mn-cs") return theme.minor.cs || theme.minor.latin || "Calibri";
  return token;
}
function fontRefOf(xml, element, ctx) {
  const style = direct2(xml, element, "style");
  const ref = style === void 0 ? void 0 : direct2(xml, style, "fontRef");
  if (ref === void 0) return NO_FONT;
  const token = attribute(ref.attrs, "idx") === "major" ? "+mj" : "+mn";
  const color = colorElement(xml, ref);
  return {
    latin: `${token}-lt`,
    ea: `${token}-ea`,
    color: color === void 0 ? void 0 : cssColor(rgbOf(xml, color, ctx))
  };
}
function bulletOf(bullet, level, counters) {
  if (bullet !== "auto") {
    counters.set(level, 0);
    return bullet;
  }
  const next = (counters.get(level) ?? 0) + 1;
  counters.set(level, next);
  const deeper = [];
  for (const key of counters.keys()) {
    if (key > level) deeper.push(key);
  }
  for (const key of deeper) counters.delete(key);
  return `${next}.`;
}
function plainOf(text) {
  if (text === void 0) return "";
  return text.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n");
}
function inlineText(xml, parent) {
  return childElements(xml, parent.openEnd, parent.closeStart).filter((item) => item.name === "t").map((item) => decodeXml(xml.slice(item.openEnd, item.closeStart))).join("");
}
function containerFill(xml, container, ctx, deck, partName, rels, groupFill) {
  const fill = childElements(xml, container.openEnd, container.closeStart).find(
    (item) => ["noFill", "solidFill", "gradFill", "blipFill", "grpFill"].includes(item.name)
  );
  if (fill === void 0) return void 0;
  return paintFillElement(xml, fill, ctx, deck, partName, rels, groupFill);
}
function paintFillElement(xml, element, ctx, deck, partName, rels, groupFill) {
  if (element.name === "noFill" || element.name === "grpFill") {
    return element.name === "grpFill" ? groupFill : NO_FILL;
  }
  if (element.name === "solidFill") {
    const color = colorElement(xml, element);
    return {
      kind: "solid",
      color: color === void 0 ? "#000000" : cssColor(rgbOf(xml, color, ctx))
    };
  }
  if (element.name === "gradFill") return gradientFill(xml, element, ctx);
  if (element.name === "blipFill") return pictureFill(xml, element, deck, partName, rels);
  return NO_FILL;
}
function fillFromRef(xml, ref, deck, partName, rels, mode) {
  const element = styleAt(deck.theme, readNumber(attribute(ref.attrs, "idx")), mode);
  if (element === void 0 || deck.theme.xml === "") return NO_FILL;
  const color = colorElement(xml, ref);
  const ctx = {
    theme: deck.theme,
    map: deck.map,
    placeholder: color === void 0 ? deck.ctx.placeholder : rgbOf(xml, color, deck.ctx)
  };
  return paintFillElement(deck.theme.xml, element, ctx, deck, partName, rels, NO_FILL);
}
function styleAt(theme, idx, mode) {
  if (idx <= 0) return void 0;
  if (mode === "bg" || idx >= 1001) {
    const index = idx >= 1001 ? idx - 1001 : idx - 1;
    return theme.bgStyles[index];
  }
  return theme.fillStyles[idx - 1];
}
function gradientFill(xml, element, ctx) {
  const list = direct2(xml, element, "gsLst");
  const stops = (list === void 0 ? [] : childElements(xml, list.openEnd, list.closeStart)).filter((item) => item.name === "gs").map((stop) => {
    const color = colorElement(xml, stop);
    return {
      offset: readNumber(attribute(stop.attrs, "pos")) / 1e5,
      color: color === void 0 ? "#000000" : cssColor(rgbOf(xml, color, ctx))
    };
  });
  const lin = direct2(xml, element, "lin");
  if (lin === void 0) return { kind: "radial", stops };
  return {
    kind: "linear",
    angle: readNumber(attribute(lin.attrs, "ang")) / 6e4,
    scaled: attribute(lin.attrs, "scaled") !== "0",
    stops
  };
}
function pictureFill(xml, element, deck, partName, rels) {
  const blip = direct2(xml, element, "blip");
  const id = blip === void 0 ? void 0 : attribute(blip.attrs, "embed", true);
  const target = id === void 0 ? void 0 : rels.get(id)?.target;
  if (target === void 0) return NO_FILL;
  const imagePath = resolvePart(partName, target);
  if (!deck.images.includes(imagePath)) deck.images.push(imagePath);
  const src = direct2(xml, element, "srcRect");
  const crop = src === void 0 ? void 0 : {
    left: readNumber(attribute(src.attrs, "l")) / 1e5,
    top: readNumber(attribute(src.attrs, "t")) / 1e5,
    right: readNumber(attribute(src.attrs, "r")) / 1e5,
    bottom: readNumber(attribute(src.attrs, "b")) / 1e5
  };
  return { kind: "picture", imagePath, crop };
}
function shapeLine(xml, spPr, style, deck) {
  if (spPr !== void 0) {
    const line = direct2(xml, spPr, "ln");
    if (line !== void 0) return lineFrom(xml, line, deck.ctx);
  }
  if (style === void 0) return void 0;
  const ref = direct2(xml, style, "lnRef");
  if (ref === void 0) return void 0;
  const themeLine = deck.theme.lineStyles[readNumber(attribute(ref.attrs, "idx")) - 1];
  if (themeLine === void 0 || deck.theme.xml === "") return void 0;
  const color = colorElement(xml, ref);
  const ctx = {
    theme: deck.theme,
    map: deck.map,
    placeholder: color === void 0 ? deck.ctx.placeholder : rgbOf(xml, color, deck.ctx)
  };
  return lineFrom(deck.theme.xml, themeLine, ctx);
}
function lineFrom(xml, line, ctx) {
  if (direct2(xml, line, "noFill") !== void 0) return void 0;
  const solid = direct2(xml, line, "solidFill");
  const colorEl = solid === void 0 ? void 0 : colorElement(xml, solid);
  const dash = direct2(xml, line, "prstDash");
  return {
    color: colorEl === void 0 ? cssColor(schemeColor("tx1", ctx)) : cssColor(rgbOf(xml, colorEl, ctx)),
    width: readNumber(attribute(line.attrs, "w"), 9525),
    dash: dashOf(dash === void 0 ? void 0 : attribute(dash.attrs, "val"))
  };
}
function dashOf(value) {
  if (value === "dot" || value === "sysDot") return "dot";
  if (value === void 0 || value === "solid") return "solid";
  return "dash";
}
function shapeShadow(xml, spPr, style, deck) {
  if (spPr !== void 0) {
    const own = direct2(xml, spPr, "effectLst");
    if (own !== void 0) return shadowFrom(xml, own, deck.ctx);
  }
  if (style === void 0) return void 0;
  const ref = direct2(xml, style, "effectRef");
  if (ref === void 0) return void 0;
  const effect = deck.theme.effectStyles[readNumber(attribute(ref.attrs, "idx")) - 1];
  if (effect === void 0 || deck.theme.xml === "") return void 0;
  const list = direct2(deck.theme.xml, effect, "effectLst");
  if (list === void 0) return void 0;
  const color = colorElement(xml, ref);
  const ctx = {
    theme: deck.theme,
    map: deck.map,
    placeholder: color === void 0 ? deck.ctx.placeholder : rgbOf(xml, color, deck.ctx)
  };
  return shadowFrom(deck.theme.xml, list, ctx);
}
function shadowFrom(xml, list, ctx) {
  const shadow = direct2(xml, list, "outerShdw");
  if (shadow === void 0) return void 0;
  const colorEl = colorElement(xml, shadow);
  const distance = readNumber(attribute(shadow.attrs, "dist"));
  const direction = readNumber(attribute(shadow.attrs, "dir")) / 6e4 * (Math.PI / 180);
  return {
    color: colorEl === void 0 ? "rgba(0, 0, 0, 0.4)" : cssColor(rgbOf(xml, colorEl, ctx)),
    blur: readNumber(attribute(shadow.attrs, "blurRad")),
    offsetX: Math.cos(direction) * distance,
    offsetY: Math.sin(direction) * distance
  };
}
function markersOf(xml) {
  const tree = findDescendant(xml, "spTree");
  if (tree === void 0) return [];
  const found = [];
  collectMarkers(xml, tree.openEnd, tree.closeStart, found);
  return found;
}
function collectMarkers(xml, from, to, found) {
  for (const element of childElements(xml, from, to)) {
    if (element.name === "grpSp") {
      collectMarkers(xml, element.openEnd, element.closeStart, found);
      continue;
    }
    const marker = markerOf(xml, element);
    if (marker !== void 0) found.push(marker);
  }
}
function markerOf(xml, element) {
  const holder = directNames(xml, element, ["nvSpPr", "nvPicPr", "nvCxnSpPr", "nvGraphicFramePr"]);
  if (holder === void 0) return void 0;
  const props = direct2(xml, holder, "nvPr");
  if (props === void 0) return void 0;
  const placeholder = direct2(xml, props, "ph");
  if (placeholder === void 0) return void 0;
  return {
    type: attribute(placeholder.attrs, "type") ?? "obj",
    idx: attribute(placeholder.attrs, "idx"),
    element
  };
}
function findPrototype(markers, wanted) {
  if (wanted.idx !== void 0) {
    const byIndex = markers.find((item) => item.idx === wanted.idx);
    if (byIndex !== void 0) return byIndex;
  }
  const kind = alias(wanted.type);
  return markers.find((item) => alias(item.type) === kind);
}
function kindsOf(markers) {
  return new Set(markers.map((marker) => alias(marker.type)));
}
function alias(type) {
  if (type === "ctrTitle") return "title";
  if (type === "subTitle" || type === "obj") return "body";
  return type;
}
function identityOf(xml, element) {
  const holder = directNames(xml, element, [
    "nvSpPr",
    "nvPicPr",
    "nvCxnSpPr",
    "nvGrpSpPr",
    "nvGraphicFramePr"
  ]);
  const node = findDescendant(
    xml,
    "cNvPr",
    (holder ?? element).openEnd,
    (holder ?? element).closeStart
  );
  if (node === void 0) return void 0;
  const id = attribute(node.attrs, "id");
  if (id === void 0) return void 0;
  return { id, name: attribute(node.attrs, "name") ?? "" };
}
function isHidden(xml, element) {
  const holder = directNames(xml, element, [
    "nvSpPr",
    "nvPicPr",
    "nvCxnSpPr",
    "nvGrpSpPr",
    "nvGraphicFramePr"
  ]);
  const node = findDescendant(
    xml,
    "cNvPr",
    (holder ?? element).openEnd,
    (holder ?? element).closeStart
  );
  return node !== void 0 && attribute(node.attrs, "hidden") === "1";
}
function direct2(xml, element, name) {
  return childElements(xml, element.openEnd, element.closeStart).find((item) => item.name === name);
}
function directNames(xml, element, names) {
  return childElements(xml, element.openEnd, element.closeStart).find(
    (item) => names.includes(item.name)
  );
}

// packages/pptx-kernel/src/zip.ts
var textEncoder = new TextEncoder();
async function readZip(archive) {
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
  const end = findEndOfCentralDirectory(archive);
  const count = view.getUint16(end + 8, true);
  let cursor = view.getUint32(end + 16, true);
  const pending = [];
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(cursor, true) !== 33639248)
      throw new Error("pptx: malformed ZIP directory");
    const method = view.getUint16(cursor + 10, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const localOffset = view.getUint32(cursor + 42, true);
    const name = new TextDecoder().decode(archive.subarray(cursor + 46, cursor + 46 + nameLength));
    const nameLengthLocal = view.getUint16(localOffset + 26, true);
    const extraLengthLocal = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + nameLengthLocal + extraLengthLocal;
    const compressed = archive.subarray(dataStart, dataStart + compressedSize);
    if (!name.endsWith("/")) {
      if (method === 0) pending.push({ name, data: Promise.resolve(compressed.slice()) });
      else if (method === 8) pending.push({ name, data: inflateRaw(compressed) });
      else throw new Error(`pptx: unsupported ZIP method ${method} for ${name}`);
    }
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  const entries = /* @__PURE__ */ new Map();
  const inflated = await Promise.all(pending.map(async (item) => item.data));
  pending.forEach((item, index) => {
    const data = inflated[index];
    if (data !== void 0) entries.set(item.name, data);
  });
  return entries;
}
async function writeZip(entries) {
  const prepared = await Promise.all(
    entries.map(async (entry) => {
      const method = entry.method === "deflate" ? 8 : 0;
      return {
        nameBytes: textEncoder.encode(entry.name),
        method,
        stored: entry.data,
        data: method === 8 ? await deflateRaw(entry.data) : entry.data,
        crc: crc32(entry.data)
      };
    })
  );
  const locals = [];
  const central = [];
  let offset = 0;
  for (const entry of prepared) {
    const { nameBytes, method, stored, data, crc } = entry;
    const local = new Uint8Array(30);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 67324752, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(8, method, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, data.byteLength, true);
    localView.setUint32(22, stored.byteLength, true);
    localView.setUint16(26, nameBytes.byteLength, true);
    locals.push(local, nameBytes, data);
    const directory = new Uint8Array(46);
    const directoryView = new DataView(directory.buffer);
    directoryView.setUint32(0, 33639248, true);
    directoryView.setUint16(4, 20, true);
    directoryView.setUint16(6, 20, true);
    directoryView.setUint16(10, method, true);
    directoryView.setUint32(16, crc, true);
    directoryView.setUint32(20, data.byteLength, true);
    directoryView.setUint32(24, stored.byteLength, true);
    directoryView.setUint16(28, nameBytes.byteLength, true);
    directoryView.setUint32(42, offset, true);
    central.push(directory, nameBytes);
    offset += local.byteLength + nameBytes.byteLength + data.byteLength;
  }
  const centralBytes = concat(central);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 101010256, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralBytes.byteLength, true);
  endView.setUint32(16, offset, true);
  return concat([...locals, centralBytes, end]);
}
function findEndOfCentralDirectory(archive) {
  const min = Math.max(0, archive.length - 22 - 65535);
  for (let index = archive.length - 22; index >= min; index -= 1) {
    if (archive[index] === 80 && archive[index + 1] === 75 && archive[index + 2] === 5 && archive[index + 3] === 6) {
      return index;
    }
  }
  throw new Error("pptx: ZIP end of central directory not found");
}
async function inflateRaw(data) {
  const stream = new Blob([toBody(data)]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function deflateRaw(data) {
  const stream = new Blob([toBody(data)]).stream().pipeThrough(new CompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
function toBody(data) {
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
}
function concat(parts) {
  const length = parts.reduce((sum, part) => sum + part.byteLength, 0);
  const out = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.byteLength;
  }
  return out;
}
var crcTable;
function crc32(bytes) {
  if (crcTable === void 0) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = (c & 1) === 1 ? 3988292384 ^ c >>> 1 : c >>> 1;
      crcTable[n] = c;
    }
  }
  let crc = 4294967295;
  for (const byte of bytes) crc = crc >>> 8 ^ crcTable[(crc ^ byte) & 255];
  return (crc ^ 4294967295) >>> 0;
}

// packages/pptx-kernel/src/pptx.ts
var textDecoder = new TextDecoder();
var textEncoder2 = new TextEncoder();
var EMU_PER_PIXEL = 914400 / 96;
var PptxDocument = class _PptxDocument {
  original;
  parts;
  order;
  edited = /* @__PURE__ */ new Map();
  scenes = /* @__PURE__ */ new Map();
  dirty = false;
  constructor(original, parts, order) {
    this.original = original;
    this.parts = parts;
    this.order = [...order];
  }
  static async open(input) {
    const bytes = input.slice();
    const parts = await readZip(bytes);
    if (!parts.has("ppt/presentation.xml")) throw new Error("pptx: presentation.xml is missing");
    return new _PptxDocument(bytes, parts, [...parts.keys()]);
  }
  slideSize() {
    const xml = this.text("ppt/presentation.xml");
    const element = findDescendant(xml, "sldSz");
    return {
      width: numberAttr(element?.attrs, "cx"),
      height: numberAttr(element?.attrs, "cy")
    };
  }
  slides() {
    const presentation = this.text("ppt/presentation.xml");
    const relations = readRelations(this.text("ppt/_rels/presentation.xml.rels"));
    const list = findDescendant(presentation, "sldIdLst");
    if (list === void 0) return [];
    const pages = [];
    for (const slideId of childElements(presentation, list.openEnd, list.closeStart)) {
      if (slideId.name !== "sldId") continue;
      const relationId = attribute(slideId.attrs, "id", true);
      const target = relationId === void 0 ? void 0 : relations.get(relationId);
      if (target === void 0 || !target.type.endsWith("/slide")) continue;
      const partName = resolvePart("ppt/presentation.xml", target.target);
      pages.push({
        partName,
        shapes: shapesOf(this.text(partName), partName, this.slideRelations(partName))
      });
    }
    return pages;
  }
  /**
   * Visual page for one slide: theme colors and fonts, master and layout chrome,
   * shape fills, and run-level text.
   */
  scene(slideIndex) {
    const cached = this.scenes.get(slideIndex);
    if (cached !== void 0) return cached;
    const built = buildSlideScene((name) => this.optionalText(name), slideIndex);
    this.scenes.set(slideIndex, built);
    return built;
  }
  /** Bytes of a package part, such as `ppt/media/image1.png`. */
  part(name) {
    const edited = this.edited.get(name);
    if (edited !== void 0) return textEncoder2.encode(edited);
    const bytes = this.parts.get(name);
    return bytes === void 0 ? void 0 : bytes.slice();
  }
  /**
   * Replace the visible text of one shape. Characters that did not change stay in their
   * original runs, so a later run such as a unit keeps its size, weight, and color.
   */
  setShapeText(slideIndex, shapeId, text) {
    this.writeShapeText(slideIndex, shapeId, (current) => distributeRuns(current, text));
  }
  /** Write one string per existing run. Run properties in the package are left untouched. */
  setShapeRuns(slideIndex, shapeId, paragraphs2) {
    this.writeShapeText(slideIndex, shapeId, () => paragraphs2.map((runs) => [...runs]));
  }
  /**
   * Append a blank slide, using the layout of `layoutFromIndex` so the deck template stays
   * in effect. Returns the new slide index.
   */
  addSlide(layoutFromIndex) {
    const existing = this.slides();
    const index = existing.length;
    const source = existing[Math.min(Math.max(layoutFromIndex, 0), Math.max(existing.length - 1, 0))];
    const layout = source === void 0 ? this.anyLayoutTarget() : this.layoutTarget(source.partName);
    this.insertSlide(index, blankSlide(), slideRelationships(layout));
    return index;
  }
  /** Insert an identical copy immediately after `index`. Returns the copy's index. */
  duplicateSlide(index) {
    const source = this.slideEntries()[index];
    if (source === void 0) throw new Error(`pptx: slide ${index} does not exist`);
    const rels = this.optionalText(slideRelsName(source.partName)) ?? slideRelationships(void 0);
    this.insertSlide(index + 1, this.text(source.partName), rels);
    return index + 1;
  }
  /**
   * Move the slide at `fromIndex` so it occupies `toIndex`.
   * Both indexes are positions in the current slide list.
   */
  moveSlide(fromIndex, toIndex) {
    const entries = this.slideEntries();
    if (entries[fromIndex] === void 0) throw new Error(`pptx: slide ${fromIndex} does not exist`);
    if (entries[toIndex] === void 0) throw new Error(`pptx: slide ${toIndex} does not exist`);
    if (fromIndex === toIndex) return;
    const next = entries.slice();
    const [moved] = next.splice(fromIndex, 1);
    if (moved === void 0) return;
    next.splice(toIndex, 0, moved);
    this.remember(
      "ppt/presentation.xml",
      replaceSlideIds(
        this.text("ppt/presentation.xml"),
        next.map((entry) => entry.tag)
      )
    );
    this.dirty = true;
    this.scenes.clear();
  }
  /** Remove one slide. A deck must keep its last slide. */
  deleteSlide(index) {
    const entries = this.slideEntries();
    const entry = entries[index];
    if (entry === void 0) throw new Error(`pptx: slide ${index} does not exist`);
    if (entries.length <= 1) throw new Error("pptx: the only slide cannot be deleted");
    this.remember(
      "ppt/presentation.xml",
      replaceSlideIds(
        this.text("ppt/presentation.xml"),
        entries.filter((item) => item !== entry).map((item) => item.tag)
      )
    );
    this.remember(
      "ppt/_rels/presentation.xml.rels",
      removeRelationship(this.text("ppt/_rels/presentation.xml.rels"), entry.relationId)
    );
    const types = this.optionalText("[Content_Types].xml");
    if (types !== void 0) {
      this.remember("[Content_Types].xml", removeContentType(types, entry.partName));
    }
    this.forget(entry.partName);
    this.forget(slideRelsName(entry.partName));
    this.dirty = true;
    this.scenes.clear();
  }
  insertSlide(index, xml, rels) {
    const file = this.nextSlideFile();
    this.remember(file.part, xml);
    this.remember(file.rels, rels);
    this.remember(
      "ppt/presentation.xml",
      insertSlideId(this.text("ppt/presentation.xml"), file.relationId, index)
    );
    this.remember(
      "ppt/_rels/presentation.xml.rels",
      insertRelationship(this.text("ppt/_rels/presentation.xml.rels"), file.relationId, file.target)
    );
    if (this.parts.has("[Content_Types].xml") || this.edited.has("[Content_Types].xml")) {
      this.remember(
        "[Content_Types].xml",
        insertContentType(this.text("[Content_Types].xml"), file.part)
      );
    }
    this.dirty = true;
    this.scenes.clear();
  }
  writeShapeText(slideIndex, shapeId, texts) {
    const slide = this.slides()[slideIndex];
    if (slide === void 0) throw new Error(`pptx: slide ${slideIndex} does not exist`);
    const xml = this.text(slide.partName);
    const shape = findShape(xml, shapeId);
    if (shape === void 0) throw new Error(`pptx: shape ${shapeId} does not exist`);
    const raw = xml.slice(shape.start, shape.end);
    const current = textGrid(raw);
    if (current.length === 0) throw new Error("pptx: shape has no text");
    const next = replaceShapeRuns(raw, texts(current));
    this.edited.set(slide.partName, xml.slice(0, shape.start) + next + xml.slice(shape.end));
    this.dirty = true;
    this.scenes.delete(slideIndex);
  }
  remember(name, xml) {
    this.edited.set(name, xml);
    if (!this.order.includes(name)) this.order.push(name);
  }
  forget(name) {
    this.edited.delete(name);
    this.parts.delete(name);
    this.order = this.order.filter((item) => item !== name);
  }
  slideEntries() {
    const presentation = this.text("ppt/presentation.xml");
    const relations = readRelations(this.text("ppt/_rels/presentation.xml.rels"));
    const list = findDescendant(presentation, "sldIdLst");
    if (list === void 0) return [];
    const entries = [];
    for (const slideId of childElements(presentation, list.openEnd, list.closeStart)) {
      if (slideId.name !== "sldId") continue;
      const relationId = attribute(slideId.attrs, "id", true);
      const target = relationId === void 0 ? void 0 : relations.get(relationId);
      if (relationId === void 0 || target === void 0 || !target.type.endsWith("/slide")) {
        continue;
      }
      entries.push({
        relationId,
        tag: presentation.slice(slideId.start, slideId.end),
        partName: resolvePart("ppt/presentation.xml", target.target)
      });
    }
    return entries;
  }
  layoutTarget(partName) {
    for (const relation of this.slideRelations(partName).values()) {
      if (relation.type.endsWith("/slideLayout")) return relation.target;
    }
    return this.anyLayoutTarget();
  }
  anyLayoutTarget() {
    for (const name of this.order) {
      const file = /^ppt\/slideLayouts\/(slideLayout\d+\.xml)$/.exec(name);
      const base = file?.[1];
      if (base !== void 0) return `../slideLayouts/${base}`;
    }
    return void 0;
  }
  nextSlideFile() {
    let max = 0;
    for (const name of /* @__PURE__ */ new Set([...this.order, ...this.edited.keys()])) {
      const match = /^ppt\/slides\/slide(\d+)\.xml$/.exec(name);
      if (match !== null) max = Math.max(max, Number(match[1]));
    }
    const number = max + 1;
    return {
      part: `ppt/slides/slide${number}.xml`,
      rels: `ppt/slides/_rels/slide${number}.xml.rels`,
      target: `slides/slide${number}.xml`,
      relationId: nextRelationId(readRelations(this.text("ppt/_rels/presentation.xml.rels")))
    };
  }
  /** Original bytes when nothing changed; otherwise a new package with untouched parts copied. */
  async save() {
    if (!this.dirty) return this.original.slice();
    return writeZip(
      this.order.map((name) => {
        const edited = this.edited.get(name);
        const data = edited === void 0 ? this.parts.get(name) : textEncoder2.encode(edited);
        if (data === void 0) throw new Error(`pptx: part ${name} disappeared`);
        return { name, data };
      })
    );
  }
  text(name) {
    const edited = this.edited.get(name);
    if (edited !== void 0) return edited;
    const bytes = this.parts.get(name);
    if (bytes === void 0) throw new Error(`pptx: part ${name} is missing`);
    return textDecoder.decode(bytes);
  }
  optionalText(name) {
    if (!this.edited.has(name) && !this.parts.has(name)) return void 0;
    return this.text(name);
  }
  slideRelations(partName) {
    const slash = partName.lastIndexOf("/");
    const rels = `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
    return this.parts.has(rels) || this.edited.has(rels) ? readRelations(this.text(rels)) : /* @__PURE__ */ new Map();
  }
};
function openPptx(input) {
  return PptxDocument.open(input);
}
function shapesOf(xml, partName, relations) {
  const tree = findDescendant(xml, "spTree");
  if (tree === void 0) return [];
  return collectShapes(xml, partName, tree.openEnd, tree.closeStart, relations, IDENTITY);
}
function collectShapes(xml, partName, from, to, relations, origin) {
  const shapes = [];
  for (const element of childElements(xml, from, to)) {
    if (element.name === "grpSp") {
      const xfrm2 = shapeXfrm(xml, element) ?? EMPTY_XFRM;
      shapes.push(
        ...collectShapes(
          xml,
          partName,
          element.openEnd,
          element.closeStart,
          relations,
          groupTransform(origin, xfrm2)
        )
      );
      continue;
    }
    if (element.name !== "sp" && element.name !== "pic") continue;
    const id = readShapeId(xml, element);
    if (id === void 0) continue;
    const xfrm = shapeXfrm(xml, element) ?? EMPTY_XFRM;
    const box = mapBox(origin, xfrm.x, xfrm.y, xfrm.width, xfrm.height);
    const text = textOf(xml.slice(element.start, element.end));
    const embed = attributeValue(xml.slice(element.start, element.end), "embed", true);
    const target = embed === void 0 ? void 0 : relations.get(embed)?.target;
    const imagePath = target === void 0 ? void 0 : resolvePart(partName, target);
    const placed = {
      id,
      name: shapeName(xml, element),
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      text
    };
    if (element.name === "pic" && imagePath !== void 0) {
      shapes.push({ kind: "picture", ...placed, imagePath });
      continue;
    }
    shapes.push({ kind: text.length > 0 ? "text" : "other", ...placed });
  }
  return shapes;
}
function findShape(xml, shapeId) {
  const tree = findDescendant(xml, "spTree");
  if (tree === void 0) return void 0;
  return findShapeIn(xml, tree.openEnd, tree.closeStart, shapeId);
}
function findShapeIn(xml, from, to, shapeId) {
  for (const element of childElements(xml, from, to)) {
    if (element.name === "grpSp") {
      const nested = findShapeIn(xml, element.openEnd, element.closeStart, shapeId);
      if (nested !== void 0) return nested;
      continue;
    }
    if ((element.name === "sp" || element.name === "pic") && readShapeId(xml, element) === shapeId) {
      return element;
    }
  }
  return void 0;
}
function readShapeId(xml, element) {
  const properties = childElements(xml, element.openEnd, element.closeStart);
  const holder = properties.find((item) => item.name === "nvSpPr" || item.name === "nvPicPr");
  const range = holder ?? element;
  const marker = findDescendant(xml, "cNvPr", range.openEnd, range.closeStart);
  return marker === void 0 ? void 0 : attribute(marker.attrs, "id");
}
function shapeName(xml, element) {
  const properties = childElements(xml, element.openEnd, element.closeStart);
  const holder = properties.find((item) => item.name === "nvSpPr" || item.name === "nvPicPr");
  const range = holder ?? element;
  const marker = findDescendant(xml, "cNvPr", range.openEnd, range.closeStart);
  return marker === void 0 ? "" : attribute(marker.attrs, "name") ?? "";
}
function textOf(raw) {
  return [...raw.matchAll(/<([A-Za-z_][\w.-]*):t(?:\s[^>]*)?>([\s\S]*?)<\/\1:t>/g)].map((match) => decodeXml(match[2] ?? "")).join("");
}
function textGrid(raw) {
  const body = findDescendant(raw, "txBody");
  if (body === void 0) return [];
  const grid = [];
  for (const paragraph of childElements(raw, body.openEnd, body.closeStart)) {
    if (paragraph.name !== "p") continue;
    const nodes = textNodesIn(raw, paragraph);
    grid.push(
      nodes.length === 0 ? [""] : nodes.map((node) => decodeXml(raw.slice(node.openEnd, node.closeStart)))
    );
  }
  return grid;
}
function replaceShapeRuns(raw, paragraphs2) {
  const body = findDescendant(raw, "txBody");
  if (body === void 0) throw new Error("pptx: shape has no text");
  const blocks = childElements(raw, body.openEnd, body.closeStart).filter(
    (item) => item.name === "p"
  );
  const edits = [];
  blocks.forEach((paragraph, index) => {
    const wanted = paragraphs2[index] ?? [""];
    const nodes = textNodesIn(raw, paragraph);
    if (nodes.length === 0) {
      const text = wanted.join("");
      if (text.length === 0) return;
      edits.push({
        start: paragraph.openEnd,
        end: paragraph.openEnd,
        xml: runXml(paragraph.prefix || "a", text)
      });
      return;
    }
    nodes.forEach((node, runIndex) => {
      const extra = runIndex === nodes.length - 1 ? wanted.slice(nodes.length).join("") : "";
      edits.push({
        start: node.start,
        end: node.end,
        xml: textElementXml(node.prefix, node.attrs, `${wanted[runIndex] ?? ""}${extra}`)
      });
    });
  });
  const overflow = paragraphs2.slice(blocks.length);
  const last = blocks[blocks.length - 1];
  if (overflow.length > 0 && last !== void 0) {
    const text = overflow.map((runs) => runs.join("")).join("\n");
    if (text.length > 0) {
      edits.push({
        start: last.closeStart,
        end: last.closeStart,
        xml: runXml(last.prefix || "a", text)
      });
    }
  }
  let next = raw;
  for (const edit of edits.toSorted((a, b) => b.start - a.start)) {
    next = next.slice(0, edit.start) + edit.xml + next.slice(edit.end);
  }
  return next;
}
function textNodesIn(xml, element) {
  const found = [];
  const walk = (from, to) => {
    for (const child of childElements(xml, from, to)) {
      if (child.name === "t") found.push(child);
      else walk(child.openEnd, child.closeStart);
    }
  };
  walk(element.openEnd, element.closeStart);
  return found;
}
function textElementXml(prefix, attrs, text) {
  const tag = prefix.length > 0 ? `${prefix}:t` : "t";
  return `<${tag}${withPreserve(attrs, text)}>${escapeXml(text)}</${tag}>`;
}
function runXml(prefix, text) {
  const run = prefix.length > 0 ? prefix : "a";
  return `<${run}:r>${textElementXml(run, "", text)}</${run}:r>`;
}
function distributeRuns(previous, nextPlain) {
  const nextLines = nextPlain.split("\n");
  if (previous.length === 0) return nextLines.map((line) => [line]);
  if (nextLines.length === previous.length) {
    return previous.map((runs, index) => distributeParagraph(runs, nextLines[index] ?? ""));
  }
  const head = previous.slice(0, -1);
  const tail = previous[previous.length - 1] ?? [""];
  return [
    ...head.map((runs, index) => distributeParagraph(runs, nextLines[index] ?? "")),
    distributeParagraph(tail, nextLines.slice(head.length).join("\n"))
  ];
}
function distributeParagraph(runs, next) {
  if (runs.length <= 1) return [next];
  const previous = runs.join("");
  let prefix = 0;
  const shared = Math.min(previous.length, next.length);
  while (prefix < shared && previous[prefix] === next[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < previous.length - prefix && suffix < next.length - prefix && previous[previous.length - 1 - suffix] === next[next.length - 1 - suffix]) {
    suffix += 1;
  }
  const owners = runOwners(runs);
  const insertAt = prefix >= previous.length ? runs.length - 1 : boundaryRun(owners, prefix);
  const buckets = runs.map(() => "");
  const push = (run, char) => {
    buckets[run] = `${buckets[run] ?? ""}${char}`;
  };
  for (let index = 0; index < prefix; index += 1) push(owners[index] ?? 0, next[index] ?? "");
  for (let index = prefix; index < next.length - suffix; index += 1)
    push(insertAt, next[index] ?? "");
  for (let index = 0; index < suffix; index += 1) {
    const oldIndex = previous.length - suffix + index;
    push(owners[oldIndex] ?? runs.length - 1, next[next.length - suffix + index] ?? "");
  }
  return buckets;
}
function runOwners(runs) {
  const owners = [];
  runs.forEach((text, index) => {
    for (let cursor = 0; cursor < text.length; cursor += 1) owners.push(index);
  });
  return owners;
}
function boundaryRun(owners, prefix) {
  const at = owners[prefix] ?? 0;
  if (prefix === 0) return at;
  const before = owners[prefix - 1] ?? at;
  return before === at ? at : before;
}
function blankSlide() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}
function slideRelsName(partName) {
  const slash = partName.lastIndexOf("/");
  return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
}
function slideRelationships(layout) {
  const relationship = layout === void 0 ? "" : `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${escapeXml(layout)}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`;
}
function replaceSlideIds(presentation, tags) {
  const list = findDescendant(presentation, "sldIdLst");
  if (list === void 0) return presentation;
  return presentation.slice(0, list.openEnd) + tags.join("") + presentation.slice(list.closeStart);
}
function removeRelationship(xml, id) {
  const root = findDescendant(xml, "Relationships");
  if (root === void 0) return xml;
  const matches = childElements(xml, root.openEnd, root.closeStart).filter(
    (item) => item.name === "Relationship" && attribute(item.attrs, "Id") === id
  );
  let next = xml;
  for (const item of matches.toReversed()) next = next.slice(0, item.start) + next.slice(item.end);
  return next;
}
function removeContentType(xml, partName) {
  const types = findDescendant(xml, "Types");
  if (types === void 0) return xml;
  const matches = childElements(xml, types.openEnd, types.closeStart).filter(
    (item) => item.name === "Override" && attribute(item.attrs, "PartName") === `/${partName}`
  );
  let next = xml;
  for (const item of matches.toReversed()) next = next.slice(0, item.start) + next.slice(item.end);
  return next;
}
function insertSlideId(presentation, relationId, index) {
  const list = findDescendant(presentation, "sldIdLst");
  const id = nextSlideId(presentation);
  const prefix = relationPrefix(presentation);
  const tag = `<p:sldId id="${id}" ${prefix}:id="${relationId}"/>`;
  if (list === void 0) {
    const root = findDescendant(presentation, "presentation");
    const at2 = root?.openEnd ?? presentation.length;
    return `${presentation.slice(0, at2)}<p:sldIdLst>${tag}</p:sldIdLst>${presentation.slice(at2)}`;
  }
  const ids = childElements(presentation, list.openEnd, list.closeStart).filter(
    (item) => item.name === "sldId"
  );
  const anchor = index <= 0 ? void 0 : ids[index - 1];
  const at = anchor?.end ?? (index <= 0 ? list.openEnd : list.closeStart);
  return presentation.slice(0, at) + tag + presentation.slice(at);
}
function nextSlideId(presentation) {
  let max = 255;
  const list = findDescendant(presentation, "sldIdLst");
  if (list === void 0) return 256;
  for (const item of childElements(presentation, list.openEnd, list.closeStart)) {
    if (item.name !== "sldId") continue;
    const id = Number(attribute(item.attrs, "id") ?? 0);
    if (Number.isFinite(id)) max = Math.max(max, id);
  }
  return max + 1;
}
function relationPrefix(presentation) {
  const list = findDescendant(presentation, "sldIdLst");
  if (list === void 0) return "r";
  for (const item of childElements(presentation, list.openEnd, list.closeStart)) {
    if (item.name !== "sldId") continue;
    const prefixed = parseAttributes(item.attrs).find(
      (attr) => attr.name === "id" && attr.prefix !== ""
    );
    if (prefixed !== void 0) return prefixed.prefix;
  }
  return "r";
}
function insertRelationship(xml, id, target) {
  const tag = `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${escapeXml(target)}"/>`;
  const close = xml.lastIndexOf("</Relationships>");
  if (close < 0) throw new Error("pptx: presentation relationships are missing");
  return xml.slice(0, close) + tag + xml.slice(close);
}
function insertContentType(xml, partName) {
  const part = `/${partName}`;
  if (xml.includes(`PartName="${part}"`)) return xml;
  const tag = `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
  const close = xml.lastIndexOf("</Types>");
  if (close < 0) return xml;
  return xml.slice(0, close) + tag + xml.slice(close);
}
function nextRelationId(relations) {
  let max = 0;
  for (const id of relations.keys()) {
    const number = Number(/^rId(\d+)$/.exec(id)?.[1] ?? 0);
    if (number > max) max = number;
  }
  return `rId${max + 1}`;
}
function withPreserve(attrs, text) {
  if (!/^\s|\s$|\n/.test(text) || attrs.includes("xml:space")) return attrs;
  return `${attrs} xml:space="preserve"`;
}
function attributeValue(raw, name, prefixed) {
  const pattern = prefixed ? new RegExp(`(?:\\s)([A-Za-z_][\\w.-]*):${name}="([^"]*)"`) : new RegExp(`(?:\\s)${name}="([^"]*)"`);
  const match = pattern.exec(raw);
  if (match === null) return void 0;
  return decodeXml(prefixed ? match[2] ?? "" : match[1] ?? "");
}
function numberAttr(attrs, name) {
  if (attrs === void 0) return 0;
  const value = attribute(attrs, name);
  const parsed = value === void 0 ? 0 : Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

// packages/pptx-kernel/web/main.ts
var page = {
  doc: null,
  slideIndex: 0,
  scale: 1,
  selectedId: void 0,
  editing: false,
  editorShape: void 0,
  thumbIndex: -1,
  images: /* @__PURE__ */ new Map(),
  saved: null,
  dirty: false,
  label: "",
  servedName: null,
  handle: null
};
var pressing = false;
var painting = false;
var plainOpen = readPlainOpen();
var menuIndex = -1;
var drag = null;
var openButton = document.querySelector("#open");
var fileInput = document.querySelector("#file");
var fileName = document.querySelector("#file-name");
var saveButton = document.querySelector("#save");
var discardButton = document.querySelector("#discard");
var addButton = document.querySelector("#add-slide");
var thumbs = document.querySelector("#thumbs");
var slideMenu = document.querySelector("#slide-menu");
var copySlideButton = document.querySelector("#copy-slide");
var deleteSlideButton = document.querySelector("#delete-slide");
var sheet = document.querySelector("#sheet");
var canvas = document.querySelector("#slide");
var editorFrame = document.querySelector("#editor-frame");
var editor = document.querySelector("#editor");
var context = canvas?.getContext("2d");
openButton?.addEventListener("click", () => {
  if (plainOpen || filePicker() === void 0 || embeddedFrame()) choosePlainFile();
  else void pickFile();
});
fileInput?.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  fileInput.value = "";
  if (file !== void 0) void openFile(file);
});
saveButton?.addEventListener("click", () => {
  void saveCurrent();
});
discardButton?.addEventListener("click", () => {
  void discardEdits();
});
addButton?.addEventListener("click", () => {
  addSlide();
});
copySlideButton?.addEventListener("click", () => {
  const index = menuIndex;
  hideMenu();
  duplicateAt(index);
});
deleteSlideButton?.addEventListener("click", () => {
  const index = menuIndex;
  hideMenu();
  deleteAt(index);
});
document.addEventListener("pointerdown", (event) => {
  if (slideMenu === null || slideMenu.hidden) return;
  if (event.target instanceof Node && slideMenu.contains(event.target)) return;
  hideMenu();
});
window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") hideMenu();
});
thumbs?.addEventListener("scroll", () => {
  hideMenu();
});
sheet?.addEventListener("pointerdown", (event) => {
  if (event.button !== 0 || canvas === null || !(event.target instanceof Node)) return;
  pressing = true;
  const box = canvas.getBoundingClientRect();
  const shape = hitTest(event.clientX - box.left, event.clientY - box.top);
  const sameText = page.editing && shape?.id === page.selectedId && editor !== null && (event.target === editor || editor.contains(event.target));
  if (sameText) return;
  event.preventDefault();
  page.selectedId = shape?.id;
  page.editing = shape?.editable === true;
  render();
  if (shape?.editable === true) placeCaret(event.clientX, event.clientY);
});
sheet?.addEventListener("mousedown", (event) => {
  if (event.button !== 0 || !(event.target instanceof Node)) return;
  if (editor !== null && (event.target === editor || editor.contains(event.target))) return;
  event.preventDefault();
});
window.addEventListener("pointerup", () => {
  if (!pressing) return;
  pressing = false;
  if (page.editing && editor !== null && document.activeElement !== editor) editor.focus();
});
window.addEventListener("pointercancel", () => {
  pressing = false;
});
canvas?.addEventListener("mousemove", (event) => {
  const shape = hitTest(event.offsetX, event.offsetY);
  canvas.style.cursor = shape?.editable === true ? "text" : "default";
});
editor?.addEventListener("input", (event) => {
  if (event.isComposing) return;
  commitEditor();
});
editor?.addEventListener("compositionend", () => {
  commitEditor();
});
editor?.addEventListener("paste", (event) => {
  event.preventDefault();
  insertPlain(event.clipboardData?.getData("text/plain") ?? "");
});
editor?.addEventListener("beforeinput", (event) => {
  if (event.inputType !== "insertParagraph" && event.inputType !== "insertLineBreak") return;
  event.preventDefault();
  insertPlain("\n");
});
editor?.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  event.preventDefault();
  page.editing = false;
  render();
  editor.blur();
});
editor?.addEventListener("blur", () => {
  window.setTimeout(() => {
    if (pressing || document.activeElement === editor) return;
    page.editing = false;
    render();
  }, 0);
});
window.addEventListener("resize", () => {
  if (page.doc !== null) render();
});
document.body.addEventListener("dragover", (event) => {
  event.preventDefault();
});
document.body.addEventListener("drop", (event) => {
  event.preventDefault();
  void openDropped(event);
});
var requested = new URLSearchParams(location.search).get("open");
if (requested !== null && safeFileName(requested)) void openRequested(requested);
async function openRequested(name) {
  try {
    const response = await fetch(encodeURIComponent(name));
    if (!response.ok) throw new Error("pptx: \u627E\u4E0D\u5230\u8FD9\u4E2A\u6587\u4EF6");
    page.handle = null;
    page.servedName = name;
    await openBytes(new Uint8Array(await response.arrayBuffer()), name);
  } catch (error) {
    note(error instanceof Error ? error.message : "\u65E0\u6CD5\u6253\u5F00\u8FD9\u4E2A\u6587\u4EF6");
  }
}
function embeddedFrame() {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}
function filePicker() {
  return window.showOpenFilePicker;
}
function readPlainOpen() {
  try {
    return sessionStorage.getItem("pptx-plain-open") === "1";
  } catch {
    return false;
  }
}
function usePlainOpen() {
  plainOpen = true;
  try {
    sessionStorage.setItem("pptx-plain-open", "1");
  } catch {
  }
}
function choosePlainFile() {
  if (fileInput === null) return;
  fileInput.value = "";
  fileInput.click();
}
function accessDenied(error) {
  if (error instanceof DOMException && (error.name === "NotAllowedError" || error.name === "SecurityError")) {
    return true;
  }
  return error instanceof Error && error.message.includes("not allowed");
}
async function pickFile() {
  const picker = filePicker();
  if (picker === void 0) {
    usePlainOpen();
    choosePlainFile();
    return;
  }
  try {
    const [handle] = await picker({
      types: [
        {
          description: "PowerPoint",
          accept: {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"]
          }
        }
      ]
    });
    if (handle === void 0) return;
    let file;
    try {
      file = await handle.getFile();
    } catch (error) {
      if (!accessDenied(error)) throw error;
      usePlainOpen();
      page.handle = null;
      choosePlainFile();
      return;
    }
    page.handle = handle;
    page.servedName = null;
    await openBytes(new Uint8Array(await file.arrayBuffer()), handle.name);
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") return;
    if (accessDenied(error)) {
      usePlainOpen();
      choosePlainFile();
      return;
    }
    note(error instanceof Error ? error.message : "\u65E0\u6CD5\u6253\u5F00\u8FD9\u4E2A\u6587\u4EF6");
  }
}
async function openDropped(event) {
  const item = event.dataTransfer?.items[0];
  const readable = item;
  const handle = await readable?.getAsFileSystemHandle?.();
  const file = item?.getAsFile() ?? event.dataTransfer?.files[0];
  if (file === void 0) return;
  page.handle = handle ?? null;
  page.servedName = null;
  await openBytes(new Uint8Array(await file.arrayBuffer()), file.name);
}
async function openFile(file) {
  page.handle = null;
  page.servedName = null;
  await openBytes(new Uint8Array(await file.arrayBuffer()), file.name);
}
async function openBytes(bytes, label) {
  const saved = bytes.slice();
  try {
    page.doc = await openPptx(saved);
  } catch (error) {
    page.doc = null;
    page.saved = null;
    clearDirty();
    if (saveButton !== null) saveButton.disabled = true;
    if (addButton !== null) addButton.disabled = true;
    note(error instanceof Error ? error.message : "\u65E0\u6CD5\u6253\u5F00\u8FD9\u4E2A\u6587\u4EF6");
    render();
    return;
  }
  page.saved = saved;
  page.label = label;
  page.slideIndex = 0;
  page.selectedId = void 0;
  page.editing = false;
  page.editorShape = void 0;
  page.thumbIndex = -1;
  page.images.clear();
  clearDirty();
  if (saveButton !== null) saveButton.disabled = false;
  if (addButton !== null) addButton.disabled = false;
  note(label);
  await loadImages();
  render();
}
async function saveCurrent() {
  const doc = page.doc;
  if (doc === null || saveButton === null || saveButton.disabled) return;
  saveButton.disabled = true;
  try {
    const bytes = await doc.save();
    await writeCurrent(bytes);
    page.saved = bytes.slice();
    const index = page.slideIndex;
    const selectedId = page.selectedId;
    page.doc = await openPptx(bytes.slice());
    page.slideIndex = Math.min(index, Math.max(page.doc.slides().length - 1, 0));
    page.selectedId = selectedId;
    page.editing = false;
    page.editorShape = void 0;
    clearDirty();
    note(page.label.length > 0 ? page.label : "\u5DF2\u4FDD\u5B58");
    await loadImages();
    render();
  } catch (error) {
    note(error instanceof Error ? error.message : "\u65E0\u6CD5\u4FDD\u5B58");
  } finally {
    if (page.doc !== null) saveButton.disabled = false;
  }
}
async function writeCurrent(bytes) {
  const body = new Blob([toBuffer(bytes)], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  });
  if (page.handle !== null) {
    const writable = await page.handle.createWritable();
    await writable.write(body);
    await writable.close();
    return;
  }
  if (page.servedName !== null) {
    const response = await fetch(`/save?name=${encodeURIComponent(page.servedName)}`, {
      method: "POST",
      body
    });
    if (!response.ok) throw new Error("\u65E0\u6CD5\u8986\u76D6\u5F53\u524D\u6587\u4EF6");
    return;
  }
  throw new Error("\u8BF7\u7528\u300C\u6253\u5F00\u300D\u9009\u62E9\u6587\u4EF6\u540E\u518D\u4FDD\u5B58");
}
async function discardEdits() {
  if (page.saved === null || !page.dirty) return;
  const index = page.slideIndex;
  page.doc = await openPptx(page.saved.slice());
  page.slideIndex = Math.min(index, Math.max(page.doc.slides().length - 1, 0));
  page.selectedId = void 0;
  page.editing = false;
  page.editorShape = void 0;
  page.thumbIndex = -1;
  page.images.clear();
  clearDirty();
  note(page.label);
  await loadImages();
  render();
}
function addSlide() {
  const doc = page.doc;
  if (doc === null) return;
  page.slideIndex = doc.addSlide(page.slideIndex);
  finishStructureChange();
}
function duplicateAt(index) {
  const doc = page.doc;
  if (doc === null || doc.slides()[index] === void 0) return;
  page.slideIndex = doc.duplicateSlide(index);
  finishStructureChange();
  void loadImages();
}
function deleteAt(index) {
  const doc = page.doc;
  if (doc === null || doc.slides().length <= 1 || doc.slides()[index] === void 0) return;
  doc.deleteSlide(index);
  page.slideIndex = Math.min(index, doc.slides().length - 1);
  finishStructureChange();
}
function finishStructureChange() {
  page.selectedId = void 0;
  page.editing = false;
  page.editorShape = void 0;
  page.thumbIndex = -1;
  if (document.activeElement === editor) editor?.blur();
  markDirty();
  render();
}
function selectSlide(index) {
  if (page.doc?.slides()[index] === void 0) return;
  page.slideIndex = index;
  page.selectedId = void 0;
  page.editing = false;
  page.editorShape = void 0;
  if (document.activeElement === editor) editor?.blur();
  render();
}
function openMenu(index, clientX, clientY) {
  if (slideMenu === null || page.doc === null) return;
  menuIndex = index;
  if (deleteSlideButton !== null) deleteSlideButton.disabled = page.doc.slides().length <= 1;
  slideMenu.hidden = false;
  const width = slideMenu.offsetWidth;
  const height = slideMenu.offsetHeight;
  const left = Math.min(clientX, Math.max(8, window.innerWidth - width - 8));
  const top = Math.min(clientY, Math.max(8, window.innerHeight - height - 8));
  slideMenu.style.left = `${left}px`;
  slideMenu.style.top = `${top}px`;
}
function hideMenu() {
  if (slideMenu === null) return;
  slideMenu.hidden = true;
  menuIndex = -1;
}
function markDirty() {
  page.dirty = true;
  if (discardButton !== null) discardButton.disabled = false;
}
function clearDirty() {
  page.dirty = false;
  if (discardButton !== null) discardButton.disabled = true;
}
function note(message) {
  if (fileName !== null) fileName.textContent = message;
}
function safeFileName(name) {
  return name.length > 0 && !name.includes("/") && !name.includes("\\") && !name.includes("..");
}
function selected() {
  return page.doc?.scene(page.slideIndex).shapes.find((shape) => shape.id === page.selectedId);
}
function syncEditor(shape) {
  if (editorFrame === null || editor === null) return;
  if (shape === void 0 || !shape.editable || !page.editing) {
    editorFrame.hidden = true;
    page.editorShape = void 0;
    return;
  }
  const switched = page.editorShape !== shape.id;
  page.editorShape = shape.id;
  editorFrame.hidden = false;
  placeEditor(shape);
  if (document.activeElement !== editor || switched) paintEditor(shape);
}
function commitEditor() {
  if (painting || editor === null) return;
  const doc = page.doc;
  const shape = selected();
  if (doc === null || shape === void 0 || !shape.editable) return;
  const grid = readEditorGrid(editor);
  const offset = grid === void 0 ? caretOffset(editor) : 0;
  try {
    if (grid === void 0)
      doc.setShapeText(page.slideIndex, shape.id, (editor.textContent ?? "").replace(/\r\n/g, "\n"));
    else doc.setShapeRuns(page.slideIndex, shape.id, grid);
  } catch (error) {
    note(error instanceof Error ? error.message : "\u65E0\u6CD5\u4FEE\u6539\u6587\u5B57");
    return;
  }
  markDirty();
  render();
  if (grid === void 0) {
    const updated = selected();
    if (updated !== void 0) paintEditor(updated);
    restoreCaret(editor, offset);
  }
}
function placeEditor(shape) {
  if (editorFrame === null || editor === null) return;
  const scale = page.scale;
  const text = shape.text;
  const run = text?.paragraphs[0]?.runs[0];
  const paragraph = text?.paragraphs[0];
  editorFrame.style.left = `${shape.x / EMU_PER_PIXEL * scale}px`;
  editorFrame.style.top = `${shape.y / EMU_PER_PIXEL * scale}px`;
  editorFrame.style.width = `${shape.width / EMU_PER_PIXEL * scale}px`;
  editorFrame.style.height = `${shape.height / EMU_PER_PIXEL * scale}px`;
  editorFrame.style.transform = `rotate(${shape.rotation}deg)`;
  const pad = text?.padding;
  editorFrame.style.padding = pad === void 0 ? "0" : `${pad.top / EMU_PER_PIXEL * scale}px ${pad.right / EMU_PER_PIXEL * scale}px ${pad.bottom / EMU_PER_PIXEL * scale}px ${pad.left / EMU_PER_PIXEL * scale}px`;
  const anchor = text?.anchor ?? "t";
  editorFrame.style.justifyContent = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
  editor.style.fontSize = "";
  editor.style.fontWeight = "";
  editor.style.fontStyle = "";
  editor.style.fontFamily = "";
  editor.style.color = "";
  editor.style.textDecoration = "";
  editor.style.textAlign = paragraph?.align === "ctr" ? "center" : paragraph?.align === "r" ? "right" : "left";
  const line = paragraph === void 0 ? (run?.fontSize ?? 18) * 96 / 72 * 1.2 : lineBox(run, paragraph);
  editor.style.lineHeight = `${line * scale}px`;
}
function paintEditor(shape) {
  const grid = (shape.text?.paragraphs ?? []).map(
    (paragraph) => paragraph.runs.map((run) => run.text)
  );
  paintGrid(shape, grid);
}
function paintGrid(shape, grid) {
  if (editor === null) return;
  const paragraphs2 = shape.text?.paragraphs ?? [];
  const nodes = [];
  painting = true;
  paragraphs2.forEach((paragraph, paragraphIndex) => {
    if (paragraphIndex > 0) nodes.push(document.createTextNode("\n"));
    paragraph.runs.forEach((run, runIndex) => {
      const span = document.createElement("span");
      span.dataset.p = String(paragraphIndex);
      span.dataset.r = String(runIndex);
      styleRun(span, run);
      span.textContent = grid[paragraphIndex]?.[runIndex] ?? run.text;
      nodes.push(span);
    });
  });
  editor.replaceChildren(...nodes);
  painting = false;
}
function styleRun(span, run) {
  const scale = page.scale;
  span.style.fontSize = `${run.fontSize * 96 / 72 * scale}px`;
  span.style.fontWeight = run.bold ? "700" : "400";
  span.style.fontStyle = run.italic ? "italic" : "normal";
  span.style.fontFamily = editorFamily(run);
  span.style.color = run.color;
  span.style.textDecoration = run.underline ? "underline" : "none";
  span.style.whiteSpace = "pre-wrap";
}
function readEditorGrid(root) {
  const grid = [];
  for (const node of root.childNodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent !== "\n") return void 0;
      continue;
    }
    if (!(node instanceof HTMLSpanElement)) return void 0;
    const paragraphIndex = Number(node.dataset.p);
    const runIndex = Number(node.dataset.r);
    if (!Number.isInteger(paragraphIndex) || !Number.isInteger(runIndex)) return void 0;
    const row = grid[paragraphIndex] ?? [];
    if (row[runIndex] !== void 0) return void 0;
    row[runIndex] = node.textContent ?? "";
    grid[paragraphIndex] = row;
  }
  for (let index = 0; index < grid.length; index += 1) {
    const row = grid[index];
    if (row === void 0) return void 0;
    for (let run = 0; run < row.length; run += 1) {
      if (row[run] === void 0) return void 0;
    }
  }
  return grid;
}
function caretOffset(root) {
  const selection = document.getSelection();
  if (selection === null || selection.rangeCount === 0 || !root.contains(selection.anchorNode))
    return 0;
  const range = selection.getRangeAt(0).cloneRange();
  range.selectNodeContents(root);
  range.setEnd(selection.getRangeAt(0).startContainer, selection.getRangeAt(0).startOffset);
  return range.toString().length;
}
function restoreCaret(root, offset) {
  const selection = document.getSelection();
  if (selection === null) return;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let remaining = offset;
  let node = walker.nextNode();
  while (node !== null) {
    const length = node.textContent?.length ?? 0;
    if (remaining <= length) {
      const range2 = document.createRange();
      range2.setStart(node, remaining);
      range2.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range2);
      return;
    }
    remaining -= length;
    node = walker.nextNode();
  }
  const range = document.createRange();
  range.selectNodeContents(root);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}
function editorFamily(run) {
  return [
    quotedFace(run.eastAsian),
    quotedFace(run.latin),
    '"Microsoft YaHei"',
    '"\u5FAE\u8F6F\u96C5\u9ED1"',
    "Calibri",
    "Arial",
    "sans-serif"
  ].filter((item) => item.length > 0).join(", ");
}
function quotedFace(name) {
  return name.length > 0 ? `"${name.replace(/"/g, "")}"` : "";
}
function placeCaret(clientX, clientY) {
  if (editor === null) return;
  editor.focus();
  void editor.offsetWidth;
  const range = caretRange(clientX, clientY);
  const selection = document.getSelection();
  if (range !== null && editor.contains(range.startContainer)) {
    range.collapse(true);
    selection?.removeAllRanges();
    selection?.addRange(range);
    return;
  }
  const end = document.createRange();
  end.selectNodeContents(editor);
  end.collapse(false);
  selection?.removeAllRanges();
  selection?.addRange(end);
}
function caretRange(clientX, clientY) {
  if (document.caretRangeFromPoint !== void 0)
    return document.caretRangeFromPoint(clientX, clientY);
  const positioned = document.caretPositionFromPoint?.(clientX, clientY);
  if (positioned === null || positioned === void 0) return null;
  const range = document.createRange();
  range.setStart(positioned.offsetNode, positioned.offset);
  range.collapse(true);
  return range;
}
function insertPlain(value) {
  if (editor === null) return;
  const selection = document.getSelection();
  if (selection === null || selection.rangeCount === 0 || !editor.contains(selection.anchorNode))
    return;
  const range = selection.getRangeAt(0);
  range.deleteContents();
  const node = document.createTextNode(value);
  range.insertNode(node);
  range.setStartAfter(node);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
  editor.dispatchEvent(new Event("input", { bubbles: true }));
}
function hitTest(x, y) {
  const shapes = page.doc?.scene(page.slideIndex).shapes ?? [];
  const px = x / page.scale;
  const py = y / page.scale;
  for (let index = shapes.length - 1; index >= 0; index -= 1) {
    const shape = shapes[index];
    if (shape === void 0 || !shape.selectable) continue;
    if (hits(shape, px, py)) return shape;
  }
  return void 0;
}
function hits(shape, px, py) {
  const x = shape.x / EMU_PER_PIXEL;
  const y = shape.y / EMU_PER_PIXEL;
  const width = shape.width / EMU_PER_PIXEL;
  const height = shape.height / EMU_PER_PIXEL;
  const dx = px - (x + width / 2);
  const dy = py - (y + height / 2);
  const rad = shape.rotation * Math.PI / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const localX = dx * cos + dy * sin;
  const localY = -dx * sin + dy * cos;
  return Math.abs(localX) <= width / 2 && Math.abs(localY) <= height / 2;
}
async function loadImages() {
  const doc = page.doc;
  if (doc === null) return;
  const paths = /* @__PURE__ */ new Set();
  doc.slides().forEach((_, index) => {
    for (const path of doc.scene(index).imagePaths) paths.add(path);
  });
  await Promise.all(
    [...paths].map(async (path) => {
      if (page.images.has(path)) return;
      const bytes = doc.part(path);
      if (bytes === void 0) return;
      try {
        page.images.set(path, await createImageBitmap(new Blob([toBuffer(bytes)])));
      } catch {
      }
    })
  );
}
function render() {
  if (canvas === null || context === null || context === void 0) return;
  const doc = page.doc;
  if (doc === null) {
    for (const button of thumbButtons()) button.remove();
    if (addButton !== null) addButton.disabled = true;
    syncEditor(void 0);
    return;
  }
  syncThumbs(doc);
  const scene = doc.scene(page.slideIndex);
  const width = Math.max(1, scene.width / EMU_PER_PIXEL);
  const height = Math.max(1, scene.height / EMU_PER_PIXEL);
  const stage = canvas.closest(".stage");
  const available = Math.max(1, (stage?.clientWidth ?? width + 48) - 48);
  const scale = Math.min(1, available / width);
  page.scale = scale;
  const ratio = window.devicePixelRatio || 1;
  paintCanvas(
    canvas,
    context,
    scene,
    width,
    height,
    scale,
    ratio,
    page.editing ? page.selectedId : void 0
  );
  const shape = selected();
  if (shape !== void 0 && shape.selectable) strokeSelection(context, shape);
  syncEditor(shape);
}
function syncThumbs(doc) {
  if (thumbs === null) return;
  const slides = doc.slides();
  const buttons = thumbButtons();
  while (buttons.length > slides.length) buttons.pop()?.remove();
  const slideChanged = page.thumbIndex !== page.slideIndex;
  slides.forEach((_, index) => {
    let button = buttons[index];
    if (button === void 0) {
      button = createThumb();
      thumbs.append(button);
      buttons.push(button);
    }
    button.dataset.index = String(index);
    button.classList.toggle("current", index === page.slideIndex);
    if (index === page.slideIndex) button.setAttribute("aria-current", "true");
    else button.removeAttribute("aria-current");
    const label = button.querySelector(".thumb-index");
    if (label !== null) label.textContent = String(index + 1);
    const preview = button.querySelector("canvas");
    if (preview !== null) paintThumb(preview, doc.scene(index));
    if (slideChanged && index === page.slideIndex) button.scrollIntoView({ block: "nearest" });
  });
  page.thumbIndex = page.slideIndex;
}
function thumbButtons() {
  if (thumbs === null) return [];
  return [...thumbs.querySelectorAll("button.thumb")];
}
function createThumb() {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "thumb";
  const label = document.createElement("span");
  label.className = "thumb-index";
  button.append(label, document.createElement("canvas"));
  button.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    const index = Number(button.dataset.index);
    if (!Number.isInteger(index)) return;
    drag = { index, pointerId: event.pointerId, startY: event.clientY, moved: false };
    button.setPointerCapture(event.pointerId);
  });
  button.addEventListener("pointermove", (event) => {
    if (drag === null || event.pointerId !== drag.pointerId) return;
    if (!drag.moved && Math.abs(event.clientY - drag.startY) < 4) return;
    drag.moved = true;
    button.classList.add("dragging");
    showDrop(dropSlot(event.clientY, drag.index));
  });
  button.addEventListener("pointerup", (event) => {
    if (drag === null || event.pointerId !== drag.pointerId) return;
    finishDrag(event.clientY);
  });
  button.addEventListener("pointercancel", (event) => {
    if (drag === null || event.pointerId !== drag.pointerId) return;
    drag = null;
    clearDrop();
  });
  button.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    const index = Number(button.dataset.index);
    if (!Number.isInteger(index)) return;
    selectSlide(index);
    openMenu(index, event.clientX, event.clientY);
  });
  return button;
}
function finishDrag(clientY) {
  const state = drag;
  drag = null;
  clearDrop();
  if (state === null) return;
  if (!state.moved) {
    selectSlide(state.index);
    return;
  }
  const to = dropSlot(clientY, state.index).to;
  const doc = page.doc;
  if (doc === null || to === state.index) {
    selectSlide(state.index);
    return;
  }
  doc.moveSlide(state.index, to);
  page.slideIndex = to;
  finishStructureChange();
}
function dropSlot(clientY, from) {
  const buttons = thumbButtons();
  let raw = buttons.length;
  for (let index = 0; index < buttons.length; index += 1) {
    const rect = buttons[index]?.getBoundingClientRect();
    if (rect !== void 0 && clientY < rect.top + rect.height / 2) {
      raw = index;
      break;
    }
  }
  return { raw, to: from < raw ? raw - 1 : raw };
}
function showDrop(slot) {
  clearDrop();
  const buttons = thumbButtons();
  const from = drag?.index;
  if (from === void 0 || slot.to === from) return;
  if (slot.raw >= buttons.length) buttons[buttons.length - 1]?.classList.add("drop-after");
  else buttons[slot.raw]?.classList.add("drop-before");
}
function clearDrop() {
  if (thumbs === null) return;
  for (const button of thumbButtons()) {
    button.classList.remove("dragging", "drop-before", "drop-after");
  }
}
function paintThumb(preview, scene) {
  const ctx = preview.getContext("2d");
  if (ctx === null) return;
  const width = Math.max(1, scene.width / EMU_PER_PIXEL);
  const height = Math.max(1, scene.height / EMU_PER_PIXEL);
  const cssWidth = 148;
  const scale = cssWidth / width;
  const ratio = window.devicePixelRatio || 1;
  paintCanvas(preview, ctx, scene, width, height, scale, ratio, void 0);
}
function paintCanvas(target, ctx, scene, width, height, scale, ratio, omitTextId) {
  target.width = Math.max(1, Math.round(width * scale * ratio));
  target.height = Math.max(1, Math.round(height * scale * ratio));
  target.style.width = `${width * scale}px`;
  target.style.height = `${height * scale}px`;
  ctx.setTransform(ratio * scale, 0, 0, ratio * scale, 0, 0);
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  paintBackdrop(ctx, scene.background, width, height);
  for (const shape of scene.shapes) drawShape(ctx, shape, omitTextId);
}
function paintBackdrop(ctx, fill, width, height) {
  if (fill.kind === "none") return;
  if (fill.kind === "picture") {
    const bitmap = page.images.get(fill.imagePath);
    if (bitmap !== void 0) drawBitmap(ctx, bitmap, fill.crop, 0, 0, width, height);
    return;
  }
  ctx.fillStyle = paintFillStyle(ctx, fill, 0, 0, width, height);
  ctx.fillRect(0, 0, width, height);
}
function drawShape(ctx, shape, omitTextId) {
  const x = shape.x / EMU_PER_PIXEL;
  const y = shape.y / EMU_PER_PIXEL;
  const width = shape.width / EMU_PER_PIXEL;
  const height = shape.height / EMU_PER_PIXEL;
  ctx.save();
  const cx = x + width / 2;
  const cy = y + height / 2;
  ctx.translate(cx, cy);
  ctx.rotate(shape.rotation * Math.PI / 180);
  if (shape.flipH || shape.flipV) ctx.scale(shape.flipH ? -1 : 1, shape.flipV ? -1 : 1);
  ctx.translate(-cx, -cy);
  if (shape.geometry === "line") {
    strokeLine(ctx, shape, x, y, width, height);
    ctx.restore();
    return;
  }
  if (shape.shadow !== void 0) {
    ctx.shadowColor = shape.shadow.color;
    ctx.shadowBlur = shape.shadow.blur / EMU_PER_PIXEL;
    ctx.shadowOffsetX = shape.shadow.offsetX / EMU_PER_PIXEL;
    ctx.shadowOffsetY = shape.shadow.offsetY / EMU_PER_PIXEL;
  }
  trace(ctx, shape, x, y, width, height);
  if (shape.fill.kind === "solid" || shape.fill.kind === "linear" || shape.fill.kind === "radial") {
    ctx.fillStyle = paintFillStyle(ctx, shape.fill, x, y, width, height);
    ctx.fill();
  } else if (shape.shadow !== void 0 && shape.fill.kind === "picture") {
    ctx.fillStyle = "#ffffff";
    ctx.fill();
  }
  ctx.shadowColor = "transparent";
  if (shape.fill.kind === "picture") {
    ctx.save();
    trace(ctx, shape, x, y, width, height);
    ctx.clip();
    const bitmap = page.images.get(shape.fill.imagePath);
    if (bitmap === void 0) {
      ctx.fillStyle = "#e6e8ea";
      ctx.fillRect(x, y, width, height);
    } else {
      drawBitmap(ctx, bitmap, shape.fill.crop, x, y, width, height);
    }
    ctx.restore();
  }
  if (shape.line !== void 0) {
    ctx.strokeStyle = shape.line.color;
    ctx.lineWidth = Math.max(shape.line.width / EMU_PER_PIXEL, 0.5);
    setDash(ctx, shape.line.dash);
    trace(ctx, shape, x, y, width, height);
    ctx.stroke();
    ctx.setLineDash([]);
  }
  if (shape.text !== void 0 && shape.id !== omitTextId) drawText(ctx, shape, x, y, width, height);
  ctx.restore();
}
function strokeSelection(ctx, shape) {
  const x = shape.x / EMU_PER_PIXEL;
  const y = shape.y / EMU_PER_PIXEL;
  const width = shape.width / EMU_PER_PIXEL;
  const height = shape.height / EMU_PER_PIXEL;
  ctx.save();
  const cx = x + width / 2;
  const cy = y + height / 2;
  ctx.translate(cx, cy);
  ctx.rotate(shape.rotation * Math.PI / 180);
  ctx.translate(-cx, -cy);
  ctx.strokeStyle = "#14a6ea";
  ctx.lineWidth = 2 / page.scale;
  ctx.setLineDash([]);
  ctx.strokeRect(x, y, width, height);
  ctx.restore();
}
function strokeLine(ctx, shape, x, y, width, height) {
  ctx.strokeStyle = shape.line?.color ?? "#000000";
  ctx.lineWidth = Math.max((shape.line?.width ?? 9525) / EMU_PER_PIXEL, 0.75);
  setDash(ctx, shape.line?.dash ?? "solid");
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(x + width, y + height);
  ctx.stroke();
  ctx.setLineDash([]);
}
function trace(ctx, shape, x, y, width, height) {
  ctx.beginPath();
  if (shape.geometry === "ellipse") {
    ctx.ellipse(
      x + width / 2,
      y + height / 2,
      Math.abs(width) / 2,
      Math.abs(height) / 2,
      0,
      0,
      Math.PI * 2
    );
    return;
  }
  if (shape.geometry === "round") {
    const radius = Math.min(Math.abs(width), Math.abs(height)) * shape.corner;
    ctx.roundRect(x, y, width, height, radius);
    return;
  }
  if (shape.geometry === "triangle") {
    ctx.moveTo(x + width / 2, y);
    ctx.lineTo(x + width, y + height);
    ctx.lineTo(x, y + height);
    ctx.closePath();
    return;
  }
  if (shape.geometry === "diamond") {
    ctx.moveTo(x + width / 2, y);
    ctx.lineTo(x + width, y + height / 2);
    ctx.lineTo(x + width / 2, y + height);
    ctx.lineTo(x, y + height / 2);
    ctx.closePath();
    return;
  }
  ctx.rect(x, y, width, height);
}
function paintFillStyle(ctx, fill, x, y, width, height) {
  if (fill.kind === "solid") return fill.color;
  if (fill.kind === "radial") {
    const radius = Math.max(width, height) / 2;
    const gradient2 = ctx.createRadialGradient(
      x + width / 2,
      y + height / 2,
      0,
      x + width / 2,
      y + height / 2,
      Math.max(1, radius)
    );
    addStops(gradient2, fill.stops);
    return gradient2;
  }
  let rad = fill.angle * Math.PI / 180;
  if (fill.scaled && width !== 0 && height !== 0) {
    rad = Math.atan2(Math.sin(rad) * height, Math.cos(rad) * width);
  }
  const dx = Math.cos(rad);
  const dy = Math.sin(rad);
  const cx = x + width / 2;
  const cy = y + height / 2;
  const half = Math.abs(width / 2 * dx) + Math.abs(height / 2 * dy);
  const gradient = ctx.createLinearGradient(
    cx - dx * half,
    cy - dy * half,
    cx + dx * half,
    cy + dy * half
  );
  addStops(gradient, fill.stops);
  return gradient;
}
function addStops(gradient, stops) {
  const ordered = [...stops].toSorted((a, b) => a.offset - b.offset);
  let last = -1;
  for (const stop of ordered) {
    let offset = Math.min(1, Math.max(0, stop.offset));
    if (offset <= last) offset = Math.min(1, last + 1e-4);
    gradient.addColorStop(offset, stop.color);
    last = offset;
  }
  if (ordered.length === 0) gradient.addColorStop(0, "#ffffff");
}
function drawBitmap(ctx, bitmap, crop, x, y, width, height) {
  const left = crop?.left ?? 0;
  const top = crop?.top ?? 0;
  const right = crop?.right ?? 0;
  const bottom = crop?.bottom ?? 0;
  const sx = bitmap.width * left;
  const sy = bitmap.height * top;
  const sw = bitmap.width * (1 - left - right);
  const sh = bitmap.height * (1 - top - bottom);
  if (sw <= 0 || sh <= 0) return;
  ctx.drawImage(bitmap, sx, sy, sw, sh, x, y, width, height);
}
function drawText(ctx, shape, x, y, width, height) {
  const text = shape.text;
  if (text === void 0) return;
  const pad = {
    left: text.padding.left / EMU_PER_PIXEL,
    top: text.padding.top / EMU_PER_PIXEL,
    right: text.padding.right / EMU_PER_PIXEL,
    bottom: text.padding.bottom / EMU_PER_PIXEL
  };
  const innerWidth = Math.max(1, width - pad.left - pad.right);
  const innerHeight = Math.max(0, height - pad.top - pad.bottom);
  ctx.save();
  trace(ctx, shape, x, y, width, height);
  ctx.clip();
  const laid = text.paragraphs.map((paragraph) => layoutParagraph(ctx, paragraph, innerWidth));
  const contentHeight = laid.reduce((sum, paragraph) => sum + paragraph.height, 0);
  let cursor = y + pad.top;
  if (text.anchor === "ctr") cursor += Math.max(0, (innerHeight - contentHeight) / 2);
  if (text.anchor === "b") cursor += Math.max(0, innerHeight - contentHeight);
  for (const paragraph of laid) {
    drawParagraph(ctx, paragraph, x + pad.left, cursor, innerWidth);
    cursor += paragraph.height;
  }
  ctx.restore();
}
function layoutParagraph(ctx, paragraph, innerWidth) {
  const marginLeft = paragraph.marginLeft / EMU_PER_PIXEL;
  const indent = paragraph.indent / EMU_PER_PIXEL;
  const contentWidth = Math.max(1, innerWidth - marginLeft);
  const lines = [];
  let line = emptyLine(paragraph);
  const push = () => {
    lines.push(line);
    line = emptyLine(paragraph);
  };
  for (const run of paragraph.runs) {
    const parts = run.text.split("\n");
    parts.forEach((part, index) => {
      if (index > 0) push();
      for (const token of tokenize(part)) {
        if (line.tokens.length === 0 && token.text.trim().length === 0) continue;
        ctx.font = fontOf(run, token.cjk);
        const measured = ctx.measureText(token.text);
        const width = measured.width;
        if (width > contentWidth && token.text.length > 1) {
          for (const char of token.text) {
            const charMeasured = ctx.measureText(char);
            const charWidth = charMeasured.width;
            if (line.width + charWidth > contentWidth && line.width > 0) push();
            addToken(line, char, charWidth, charMeasured, token.cjk, run, paragraph);
          }
          continue;
        }
        if (line.width + width > contentWidth && line.width > 0) push();
        addToken(line, token.text, width, measured, token.cjk, run, paragraph);
      }
    });
  }
  if (line.tokens.length > 0 || lines.length === 0) lines.push(line);
  const spaceBefore = paragraph.spaceBefore * 96 / 72;
  const spaceAfter = paragraph.spaceAfter * 96 / 72;
  const height = spaceBefore + spaceAfter + lines.reduce((sum, item) => sum + item.height, 0);
  return {
    align: paragraph.align,
    bullet: paragraph.bullet,
    bulletRun: paragraph.runs[0],
    marginLeft,
    indent,
    spaceBefore,
    spaceAfter,
    lines,
    height
  };
}
function emptyLine(paragraph) {
  return { width: 0, height: lineBox(paragraph.runs[0], paragraph), tokens: [] };
}
function addToken(line, text, width, measured, cjk, run, paragraph) {
  const size = run.fontSize * 96 / 72;
  const { ascent, descent } = fontEdges(measured, size);
  line.tokens.push({ text, width, ascent, descent, cjk, run });
  line.width += width;
  line.height = Math.max(line.height, lineBox(run, paragraph));
}
function fontEdges(measured, size) {
  const ascent = measured.fontBoundingBoxAscent;
  const descent = measured.fontBoundingBoxDescent;
  if (ascent > 0 || descent > 0) return { ascent, descent };
  return { ascent: size * 0.8, descent: size * 0.2 };
}
function lineBaseline(top, height, ascent, descent) {
  return top + (height - (ascent + descent)) / 2 + ascent;
}
function lineBox(run, paragraph) {
  if (paragraph.linePt > 0) return paragraph.linePt * 96 / 72;
  const size = (run?.fontSize ?? 18) * 96 / 72;
  return size * 1.2 * paragraph.linePct;
}
function drawParagraph(ctx, paragraph, x, y, innerWidth) {
  let cursor = y + paragraph.spaceBefore;
  const textX = x + paragraph.marginLeft;
  paragraph.lines.forEach((line, index) => {
    let offset = textX;
    if (paragraph.align === "ctr") offset = x + (innerWidth - line.width) / 2;
    if (paragraph.align === "r") offset = x + innerWidth - line.width;
    ctx.textBaseline = "alphabetic";
    const baseline = sharedBaseline(ctx, paragraph, line, cursor, index);
    if (index === 0 && paragraph.bullet !== void 0 && paragraph.bulletRun !== void 0) {
      ctx.font = fontOf(paragraph.bulletRun, isCjk(paragraph.bullet));
      ctx.fillStyle = paragraph.bulletRun.color;
      ctx.fillText(paragraph.bullet, textX + paragraph.indent, baseline);
    }
    for (const token of line.tokens) {
      ctx.font = fontOf(token.run, token.cjk);
      ctx.fillStyle = token.run.color;
      ctx.fillText(token.text, offset, baseline);
      if (token.run.underline) {
        const size = token.run.fontSize * 96 / 72;
        const underline = baseline + size * 0.12;
        ctx.strokeStyle = token.run.color;
        ctx.lineWidth = Math.max(1, size / 14);
        ctx.beginPath();
        ctx.moveTo(offset, underline);
        ctx.lineTo(offset + token.width, underline);
        ctx.stroke();
      }
      offset += token.width;
    }
    cursor += line.height;
  });
}
function sharedBaseline(ctx, paragraph, line, cursor, index) {
  let ascent = 0;
  let descent = 0;
  for (const token of line.tokens) {
    ascent = Math.max(ascent, token.ascent);
    descent = Math.max(descent, token.descent);
  }
  if (index === 0 && paragraph.bullet !== void 0 && paragraph.bulletRun !== void 0) {
    ctx.font = fontOf(paragraph.bulletRun, isCjk(paragraph.bullet));
    const size = paragraph.bulletRun.fontSize * 96 / 72;
    const edge = fontEdges(ctx.measureText(paragraph.bullet), size);
    ascent = Math.max(ascent, edge.ascent);
    descent = Math.max(descent, edge.descent);
  }
  if (ascent === 0 && descent === 0) return cursor + line.height * 0.8;
  return lineBaseline(cursor, line.height, ascent, descent);
}
function fontOf(run, cjk) {
  const family = cjk ? run.eastAsian : run.latin;
  const size = run.fontSize * 96 / 72;
  const quoted = family.length > 0 ? `"${family.replace(/"/g, "")}"` : "";
  const fallback = cjk ? '"Microsoft YaHei", "\u5FAE\u8F6F\u96C5\u9ED1", "PingFang SC", "Noto Sans SC", sans-serif' : 'Calibri, "Segoe UI", Arial, sans-serif';
  const stack = [quoted, fallback].filter((item) => item.length > 0).join(", ");
  return [run.italic ? "italic" : "", run.bold ? "700" : "400", `${size}px`, stack].filter((item) => item.length > 0).join(" ");
}
function tokenize(text) {
  const tokens = [];
  for (const match of text.matchAll(
    /[\u3400-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]|[^\s\u3400-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]+|\s+/gu
  )) {
    const value = match[0] ?? "";
    tokens.push({ text: value, cjk: isCjk(value) });
  }
  return tokens;
}
function isCjk(text) {
  const code = text.codePointAt(0) ?? 0;
  return code >= 13312 && code <= 40959 || code >= 63744 && code <= 64255 || code >= 12352 && code <= 12543 || code >= 44032 && code <= 55215;
}
function setDash(ctx, dash) {
  if (dash === "dot") ctx.setLineDash([1.5, 3]);
  else if (dash === "dash") ctx.setLineDash([8, 4]);
  else ctx.setLineDash([]);
}
function toBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}
