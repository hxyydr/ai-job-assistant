/**
 * 零依赖中文 PDF 生成器。
 * 把轻量 Markdown 文本渲染为可被 PDF 阅读器打开、文字可选可搜索的 PDF。
 * 做法：解析系统 TrueType 字体（黑体），以 CIDFontType2 + Identity-H 编码嵌入，
 * 手写 PDF 对象、内容流、ToUnicode CMap 与交叉引用表。不引入任何第三方包。
 */
'use strict';

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

// ---------- 字体候选（按优先级） ----------
// 首选项目内预生成的黑体子集（约 2MB，覆盖 GB2312 常用字），大幅减小 PDF 体积；
// 找不到时回退到系统完整黑体，保证零依赖与跨平台可用性。
const FONT_CANDIDATES = [
  path.join(__dirname, '..', 'public', 'fonts', 'simhei-subset.ttf'),
  'C:\\Windows\\Fonts\\simhei.ttf',
  '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
  '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
];

let _fontCache = null;      // 子集字体缓存
let _fullFontCache = null;  // 系统完整字体缓存（子集缺字时回退）
let _styleCache = {};       // 按字体文件路径缓存解析结果

// 当前导出样式（由 generatePdf 注入）：默认用原生蓝色模板；带 source 时套用原文件样式
let ACTIVE = null;
function styleBodySize() { return (ACTIVE && ACTIVE.fontSizePt) ? ACTIVE.fontSizePt : 10.5; }
function styleHeadingColor() { return (ACTIVE && ACTIVE.headingColor) ? ACTIVE.headingColor : '1F4E9C'; }
// hex -> "r g b rg"
function hexRg(hex, fallback) {
  const h = hex || fallback;
  if (!/^[0-9A-Fa-f]{6}$/.test(h)) return null;
  const r = parseInt(h.slice(0, 2), 16) / 255, g = parseInt(h.slice(2, 4), 16) / 255, b = parseInt(h.slice(4, 6), 16) / 255;
  return r.toFixed(2) + ' ' + g.toFixed(2) + ' ' + b.toFixed(2) + ' rg';
}
/**
 * 把 .ttc（字体集合）的某个 face 扁平化成独立 .ttf 字节。
 * 原因：TTC 表目录里的 offset 是「相对整个 TTC 文件」的绝对偏移；PDF 的
 * FontFile2 必须是独立 sfnt，直接 subarray(base) 嵌入会让表偏移全部错位，
 * 导致中文渲染成空白。这里重写 sfnt 头与表目录，把表数据重排为相对新文件的偏移。
 */
function flattenTtc(buf, faceIndex) {
  if (buf.length < 12 || buf.readUInt32BE(0) !== 0x74746366) return buf; // 非 TTC 原样返回
  const numFaces = buf.readUInt32BE(8);
  const fi = (faceIndex >= 0 && faceIndex < numFaces) ? faceIndex : 0;
  const faceOff = buf.readUInt32BE(12 + fi * 4);
  const numTables = buf.readUInt16BE(faceOff + 4);
  const headerLen = 12, dirLen = numTables * 16;
  let dataOff = headerLen + dirLen;
  const tables = [];
  for (let i = 0; i < numTables; i++) {
    const o = faceOff + 12 + i * 16;
    const tag = buf.toString('latin1', o, o + 4);
    const checksum = buf.readUInt32BE(o + 4);
    const srcOff = buf.readUInt32BE(o + 8);
    const len = buf.readUInt32BE(o + 12);
    tables.push({ tag, checksum, srcOff, len, newOff: dataOff });
    dataOff += Math.ceil(len / 4) * 4; // 4 字节对齐
  }
  const out = Buffer.alloc(dataOff);
  out.writeUInt32BE(0x00010000, 0);      // sfntVersion: TrueType
  out.writeUInt16BE(numTables, 4);
  let p = 1; while (p * 2 <= numTables) p *= 2;
  out.writeUInt16BE(p * 16, 6);                       // searchRange
  out.writeUInt16BE(Math.log2(p), 8);                 // entrySelector
  out.writeUInt16BE(numTables * 16 - p * 16, 10);     // rangeShift
  tables.forEach((t, i) => {
    const o = headerLen + i * 16;
    out.write(t.tag, o, 4, 'latin1');
    out.writeUInt32BE(t.checksum, o + 4);
    out.writeUInt32BE(t.newOff, o + 8);
    out.writeUInt32BE(t.len, o + 12);
    buf.copy(out, t.newOff, t.srcOff, t.srcOff + t.len);
  });
  return out;
}

/**
 * 字体子集化：把完整字体重建为「只含文本实际用到字形」的极小字体。
 * 关键：重新把 GID 映射为紧凑编号（0..N），并同步重建 glyf/loca/hmtx/cmap/maxp/hhea。
 * 这样即便原字体是 17MB 的宋体 TTC，导出 PDF 也只嵌入几十 KB，生成从 2s 降到几十 ms。
 * 复合字形（composite glyph）内部引用的组件 GID 也会一并重写，避免中文缺字。
 * 返回 {buf, cmap:{map,maxGid}, widths, numGlyphs}；失败时回退原字体对象。
 */
function buildCmap4(codepoints, oldMap, old2new) {
  var segs = [];
  codepoints.forEach(function (c) {
    var og = oldMap[c];
    if (!og) return;
    var ng = old2new[og];
    if (ng == null) return;
    segs.push({ start: c, end: c, delta: (ng - c) & 0xFFFF });
  });
  segs.push({ start: 0xFFFF, end: 0xFFFF, delta: 1 }); // 必需的终止段
  var n = segs.length;
  var subLen = 16 + n * 8;
  var sub = Buffer.alloc(subLen);
  sub.writeUInt16BE(4, 0); sub.writeUInt16BE(subLen, 2); sub.writeUInt16BE(0, 4);
  var p = 1; while (p * 2 <= n) p *= 2;
  sub.writeUInt16BE(n * 2, 6); sub.writeUInt16BE(Math.log2(p), 8); sub.writeUInt16BE(n * 2 - p * 2, 10);
  var eo = 14, so = eo + n * 2 + 2, dO = so + n * 2, ro = dO + n * 2;
  segs.forEach(function (s, i) {
    sub.writeUInt16BE(s.end, eo + i * 2);
    sub.writeUInt16BE(s.start, so + i * 2);
    sub.writeUInt16BE(s.delta & 0xFFFF, dO + i * 2); // idDelta 以无符号位模式写入（读取端按补码还原）
    sub.writeUInt16BE(0, ro + i * 2); // idRangeOffset=0 → 用 idDelta
  });
  var table = Buffer.alloc(12 + subLen);
  table.writeUInt16BE(0, 0); table.writeUInt16BE(1, 2);
  table.writeUInt16BE(3, 4); table.writeUInt16BE(1, 6); table.writeUInt32BE(12, 8);
  sub.copy(table, 12);
  return table;
}
function subsetFontForText(font, text) {
  try {
    var buf = font.buf, T = font.tables;
    var glyfOff = T.glyf.off, locaOff = T.loca.off, headOff = T.head.off, hheaOff = T.hhea.off, hmtxOff = T.hmtx.off;
    var ilf = buf.readInt16BE(headOff + 50);
    var numGlyphs = buf.readUInt16BE(T.maxp.off + 4);
    var locaAt = function (g) { return ilf === 0 ? buf.readUInt16BE(locaOff + g * 2) * 2 : buf.readUInt32BE(locaOff + g * 4); };
    var numHM = buf.readUInt16BE(hheaOff + 34);
    var advOf = function (g) { return g < numHM ? buf.readUInt16BE(hmtxOff + g * 4) : buf.readUInt16BE(hmtxOff + (numHM - 1) * 4); };
    var lsbOf = function (g) { return g < numHM ? buf.readInt16BE(hmtxOff + g * 4 + 2) : buf.readInt16BE(hmtxOff + numHM * 4 + (g - numHM) * 2); };

    // 1) 收集用到的字符与初始 GID
    var cps = new Set();
    for (var i = 0; i < String(text).length; i++) {
      var cp = String(text).charCodeAt(i);
      if (cp >= 0x20) cps.add(cp);
    }
    var oldSet = new Set([0]);
    cps.forEach(function (c) { var g = font.cmap.map[c]; if (g) oldSet.add(g); });
    // 2) 展开复合字形依赖的组件 GID
    var queue = Array.from(oldSet);
    while (queue.length) {
      var g = queue.pop();
      var s = locaAt(g), e = locaAt(g + 1);
      if (e <= s || s + 10 > buf.length) continue;
      if (buf.readInt16BE(glyfOff + s) >= 0) continue; // 简单字形
      var p = glyfOff + s + 10, fl;
      do {
        if (p + 4 > buf.length) break;
        fl = buf.readUInt16BE(p); p += 2;
        var cg = buf.readUInt16BE(p); p += 2;
        if (!oldSet.has(cg)) { oldSet.add(cg); queue.push(cg); }
        if (fl & 0x0001) p += 4; else p += 2;          // ARG_1_AND_2_ARE_WORDS
        if (fl & 0x0008) p += 2;                        // WE_HAVE_A_SCALE
        else if (fl & 0x0040) p += 4;                   // X_AND_Y_SCALE
        else if (fl & 0x0080) p += 8;                   // TWO_BY_TWO
      } while (fl & 0x0020);                            // MORE_COMPONENTS
    }
    var oldGids = Array.from(oldSet).sort(function (a, b) { return a - b; });
    var old2new = Object.create(null);
    oldGids.forEach(function (og, i) { old2new[og] = i; });
    var newNum = oldGids.length;

    // 3) 重建 glyf（复合字形的组件 GID 改写为新编号）
    var parts = [], off = 0, newLoca = new Uint32Array(newNum + 1);
    oldGids.forEach(function (og, ng) {
      newLoca[ng] = off;
      var s = locaAt(og), e = locaAt(og + 1);
      if (e <= s || s < 0 || e > buf.length) return;
      var data = Buffer.from(buf.subarray(glyfOff + s, glyfOff + e));
      if (data.length >= 12 && data.readInt16BE(0) < 0) {
        var q = 10, flg;
        do {
          if (q + 4 > data.length) break;
          flg = data.readUInt16BE(q); q += 2;
          var compOld = data.readUInt16BE(q);
          data.writeUInt16BE(old2new[compOld] == null ? 0 : old2new[compOld], q); q += 2;
          if (flg & 0x0001) q += 4; else q += 2;
          if (flg & 0x0008) q += 2; else if (flg & 0x0040) q += 4; else if (flg & 0x0080) q += 8;
        } while (flg & 0x0020);
      }
      parts.push(data); off += data.length;
      while (off % 4) { parts.push(Buffer.from([0])); off++; }
    });
    newLoca[newNum] = off;
    var newGlyf = Buffer.concat(parts);
    var newLocaBuf = Buffer.alloc((newNum + 1) * 4);
    for (var k = 0; k <= newNum; k++) newLocaBuf.writeUInt32BE(newLoca[k], k * 4);

    // 4) 重建 hmtx / hhea
    var newHmtx = Buffer.alloc(newNum * 4);
    oldGids.forEach(function (og, i) { newHmtx.writeUInt16BE(advOf(og), i * 4); newHmtx.writeInt16BE(lsbOf(og), i * 4 + 2); });
    var newHhea = Buffer.from(buf.subarray(hheaOff, hheaOff + 36));
    newHhea.writeUInt16BE(newNum, 34); // numberOfHMetrics

    // 5) 重建 cmap / maxp / head
    var newCmap = buildCmap4(cps, font.cmap.map, old2new);
    var newMaxp = Buffer.from(buf.subarray(T.maxp.off, T.maxp.off + 32));
    newMaxp.writeUInt16BE(newNum, 4);
    var newHead = Buffer.from(buf.subarray(headOff, headOff + 54));
    newHead.writeInt16BE(1, 50);          // indexToLocFormat=1 (long)
    newHead.writeUInt32BE(0, 8);          // checkSumAdjustment=0（PDF 不校验）
    var newName = Buffer.from(buf.subarray(T.name.off, T.name.off + T.name.len));
    var newPost = Buffer.from(buf.subarray(T.post.off, T.post.off + T.post.len));
    var newOS2 = T['OS/2'] ? Buffer.from(buf.subarray(T['OS/2'].off, T['OS/2'].off + T['OS/2'].len)) : null;

    // 6) 只组装必要表，丢弃 EBDT/EBLC/vmtx/GSUB 等巨表
    var rows = [
      { tag: 'OS/2', data: newOS2 }, { tag: 'cmap', data: newCmap }, { tag: 'glyf', data: newGlyf },
      { tag: 'head', data: newHead }, { tag: 'hhea', data: newHhea }, { tag: 'hmtx', data: newHmtx },
      { tag: 'loca', data: newLocaBuf }, { tag: 'maxp', data: newMaxp }, { tag: 'name', data: newName },
      { tag: 'post', data: newPost }
    ].filter(function (r) { return r.data && r.data.length; });
    rows.sort(function (a, b) { return a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0; });
    var dirOff = 12 + rows.length * 16, dataOff = dirOff;
    rows.forEach(function (r) { r.no = dataOff; dataOff += Math.ceil(r.data.length / 4) * 4; });
    var out = Buffer.alloc(dataOff);
    out.writeUInt32BE(0x00010000, 0); out.writeUInt16BE(rows.length, 4);
    var sp = 1; while (sp * 2 <= rows.length) sp *= 2;
    out.writeUInt16BE(sp * 16, 6); out.writeUInt16BE(Math.log2(sp), 8); out.writeUInt16BE(rows.length * 16 - sp * 16, 10);
    rows.forEach(function (r, i) {
      var o = 12 + i * 16;
      out.write(r.tag, o, 4, 'ascii'); out.writeUInt32BE(0, o + 4);
      out.writeUInt32BE(r.no, o + 8); out.writeUInt32BE(r.data.length, o + 12);
      r.data.copy(out, r.no);
    });

    // 7) 更新映射：新 GID 与紧凑宽度
    var newMap = Object.create(null), newWidths = new Array(newNum).fill(500);
    cps.forEach(function (c) { var ng = old2new[font.cmap.map[c]]; if (ng != null) newMap[c] = ng; });
    oldGids.forEach(function (og, i) { newWidths[i] = font.widths[og] || 500; });
    return {
      buf: out, base: 0, tables: null,
      cmap: { map: newMap, maxGid: Math.max(0, newNum - 1) },
      widths: newWidths, numGlyphs: newNum,
      unitsPerEm: font.unitsPerEm, ascender: font.ascender, descender: font.descender, fontBBox: font.fontBBox,
      subsetted: true
    };
  } catch (e) { if (global.__PDF_DEBUG__) console.error('[subset fail]', e && e.stack || e); return font; } // 子集化失败则回退完整字体，保证可用
}

function parseFontFile(p) {
  let buf;
  try { buf = fs.readFileSync(p); } catch (e) { return null; }
  // 先把 TTC 扁平化成独立 TTF：之后所有表偏移都是相对 buf 起点，嵌入也不会错位
  buf = flattenTtc(buf, 0);
  let base = 0;
  const tables = {};
  const numTables = buf.readUInt16BE(base + 4);
  for (let i = 0; i < numTables; i++) {
    const o = base + 12 + i * 16;
    const tag = buf.toString('ascii', o, o + 4);
    tables[tag] = { off: buf.readUInt32BE(o + 8), len: buf.readUInt32BE(o + 12) };
  }
  for (const need of ['cmap', 'head', 'hhea', 'hmtx', 'maxp', 'post']) {
    if (!tables[need]) return null;
  }
  const head = tables.head.off;
  const unitsPerEm = buf.readUInt16BE(head + 18) || 1000;
  const hhea = tables.hhea.off;
  const numHMetrics = buf.readUInt16BE(hhea + 34);
  const cmap = parseCmap(buf, tables.cmap.off);
  const widths = parseWidths(buf, tables.hmtx.off, numHMetrics, cmap.maxGid, unitsPerEm);
  return {
    buf, base, tables, cmap, widths, unitsPerEm, numHMetrics, path: p,
    ascender: buf.readInt16BE(hhea + 4),
    descender: buf.readInt16BE(hhea + 6),
    fontBBox: [
      buf.readInt16BE(head + 36), buf.readInt16BE(head + 38),
      buf.readInt16BE(head + 40), buf.readInt16BE(head + 42)
    ],
    numGlyphs: buf.readUInt16BE(tables.maxp.off + 4)
  };
}
function loadFont() {
  if (_fontCache) return _fontCache;
  for (const p of FONT_CANDIDATES) {
    if (!fs.existsSync(p)) continue;
    const f = parseFontFile(p);
    if (f) { _fontCache = f; return f; }
  }
  throw new Error('未找到可用的中文字体（需 simhei.ttf 等）');
}
function loadFullFont() {
  if (_fullFontCache) return _fullFontCache;
  // 跳过第一个子集字体候选，尝试系统完整字体
  for (const p of FONT_CANDIDATES.slice(1)) {
    if (!fs.existsSync(p)) continue;
    const f = parseFontFile(p);
    if (f) { _fullFontCache = f; return f; }
  }
  return null;
}
// 加载指定路径字体（原文件字体族），带缓存；失败返回 null
function loadFontPath(p) {
  if (!p || !fs.existsSync(p)) return null;
  if (_styleCache[p]) return _styleCache[p];
  const f = parseFontFile(p);
  if (f) _styleCache[p] = f;
  return f;
}
// 选字体：若指定了原文件字体路径则用之（TTC 已由 flattenTtc 扁平化，映射不再错位）；
// 否则用项目内置黑体子集。这样导出 PDF 能真正套用原简历字体。
function ensureFontFor(text) {
  if (ACTIVE && ACTIVE.fontPath) {
    const f = loadFontPath(ACTIVE.fontPath);
    if (f) return f;
  }
  let f = loadFont();
  let missing = 0;
  for (const ch of String(text || '')) {
    const code = ch.codePointAt(0);
    if (code < 0x20) continue;
    if (!f.cmap.map[code]) missing++;
  }
  if (missing > 0) {
    const full = loadFullFont();
    if (full) return full;
  }
  return f;
}

// cmap format 4 → {map: {unicode: gid}, maxGid}
function parseCmap(buf, off) {
  const n = buf.readUInt16BE(off + 2);
  let sub = null;
  // 优先 Windows Unicode BMP (platform 3, enc 1)，其次 (3,0)
  const order = [];
  for (let i = 0; i < n; i++) {
    const so = off + 4 + i * 8;
    const pid = buf.readUInt16BE(so), eid = buf.readUInt16BE(so + 2);
    const o = off + buf.readUInt32BE(so + 4);
    order.push({ pid, eid, o });
  }
  order.sort((a, b) => {
    const score = x => (x.pid === 3 && x.eid === 1) ? 3 : (x.pid === 3 && x.eid === 0) ? 2 : (x.pid === 0) ? 1 : 0;
    return score(b) - score(a);
  });
  for (const it of order) {
    const fmt = buf.readUInt16BE(it.o);
    if (fmt === 4) { sub = it.o; break; }
  }
  if (sub === null) throw new Error('字体 cmap 无 format 4 子表');
  const segCountX2 = buf.readUInt16BE(sub + 6);
  const seg = segCountX2 / 2;
  const endBase = sub + 14;
  const startBase = endBase + segCountX2 + 2;
  const deltaBase = startBase + segCountX2;
  const rangeBase = deltaBase + segCountX2;
  const map = Object.create(null);
  let maxGid = 0;
  for (let i = 0; i < seg; i++) {
    const end = buf.readUInt16BE(endBase + i * 2);
    const start = buf.readUInt16BE(startBase + i * 2);
    const delta = buf.readInt16BE(deltaBase + i * 2);
    const rangeOff = buf.readUInt16BE(rangeBase + i * 2);
    for (let c = start; c <= end && c !== 0xFFFF; c++) {
      let gid;
      if (rangeOff === 0) {
        gid = (c + delta) & 0xFFFF;
      } else {
        const addr = rangeBase + i * 2 + rangeOff + (c - start) * 2;
        if (addr + 1 >= buf.length) continue;
        const g = buf.readUInt16BE(addr);
        gid = g === 0 ? 0 : (g + delta) & 0xFFFF;
      }
      if (gid > 0) { map[c] = gid; if (gid > maxGid) maxGid = gid; }
    }
  }
  return { map, maxGid };
}

function parseWidths(buf, hmtxOff, numHMetrics, maxGid, unitsPerEm) {
  const widths = [];
  let last = 0;
  for (let gid = 0; gid <= maxGid; gid++) {
    let w;
    if (gid < numHMetrics) {
      w = buf.readUInt16BE(hmtxOff + gid * 4);
      last = w;
    } else {
      w = last;
    }
    widths[gid] = Math.round(w * 1000 / unitsPerEm);
  }
  return widths;
}

// ---------- Markdown → 排版行 ----------
function mdToLines(src) {
  const lines = String(src == null ? '' : src).replace(/\r/g, '').split('\n');
  const out = [];
  lines.forEach(line => {
    const t = line.replace(/\s+$/, '');
    if (!t.trim()) { out.push({ text: '', size: 0, gap: 6 }); return; }
    if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(t)) {
      out.push({ rule: true });
      return;
    }
    const bs = styleBodySize();
    const h = t.match(/^\s*(#{1,4})\s+(.*)$/);
    if (h) {
      const lv = h[1].length;
      out.push({ text: stripInline(h[2]), size: bs * (lv === 1 ? 1.7 : lv === 2 ? 1.4 : 1.25), bold: true, color: styleHeadingColor(), before: 12, after: 5 });
      return;
    }
    const ol = t.match(/^\s*(\d+)[.)]\s+(.*)$/);
    if (ol) { out.push({ text: ol[1] + '. ' + stripInline(ol[2]), size: bs, indent: 14 }); return; }
    const ul = t.match(/^\s*[-*•·]\s+(.*)$/);
    // 用「·」而非「•」：内置黑体子集 cmap 无 U+2022(•)，渲染会变方框；U+00B7(·) 有字形
    if (ul) { out.push({ text: '· ' + stripInline(ul[1]), size: bs, indent: 14 }); return; }
    out.push({ text: stripInline(t), size: bs });
  });
  return out;
}
function stripInline(s) {
  return String(s).replace(/\*\*([^*]+)\*\*/g, '$1').replace(/`([^`]+)`/g, '$1').replace(/\[([^\]]+)\]\([^)]*\)/g, '$1');
}

// ---------- 文本按宽度拆行 ----------
function wrapText(text, font, maxWidth) {
  const chars = Array.from(text);
  const lines = [];
  let cur = '', w = 0;
  for (const ch of chars) {
    const cw = charWidth(ch, font);
    if (w + cw > maxWidth && cur) { lines.push(cur); cur = ''; w = 0; }
    cur += ch; w += cw;
  }
  lines.push(cur);
  return lines;
}
function charWidth(ch, font) {
  const code = ch.codePointAt(0);
  const gid = font.cmap.map[code] || 0;
  const u = gid ? (font.widths[gid] || 500) : 500;
  return u / 1000; // 以 em 为单位
}

// ---------- PDF 工具 ----------
function pdfEscape(s) { return String(s).replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)'); }
function hex4(n) { return ('0000' + n.toString(16)).slice(-4).toUpperCase(); }
function textToHex(text, font) {
  let h = '';
  for (const ch of text) {
    const code = ch.codePointAt(0);
    let gid = font.cmap.map[code] || 0;
    h += hex4(gid);
  }
  return h;
}

/**
 * 生成 PDF Buffer
 * @param {string} text 轻量 Markdown
 * @param {string} title 标题（仅作文档信息）
 * @returns {Buffer}
 */
/** 解析 JPEG 宽高（SOF 标记），用于把证件照作为 Image XObject 嵌入 */
function jpegSize(buf) {
  try {
    let i = 2;
    while (i < buf.length - 9) {
      if (buf[i] !== 0xFF) { i++; continue; }
      const m = buf[i + 1];
      if (m === 0xD8 || m === 0xD9 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
      const len = buf.readUInt16BE(i + 2);
      const isSof = (m >= 0xC0 && m <= 0xC3) || (m >= 0xC5 && m <= 0xC7) ||
                    (m >= 0xC9 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF);
      if (isSof) return { w: buf.readUInt16BE(i + 7), h: buf.readUInt16BE(i + 5) };
      i += 2 + len;
    }
  } catch (e) {}
  return null;
}

/**
 * PNG -> PDF Image XObject：完整解码 PNG（inflate + 逐行反滤波），剥离 alpha 合成到白底，
 * 输出纯 RGB 后以 DCTDecode 无关的方式重新 FlateDecode 嵌入。
 * 关键：不能把带 alpha 的 RGBA 流直接塞进 PDF 的 DeviceRGB+Predictor —— 通道数不匹配
 * 会让 PDF 阅读器解析失败（对象变损坏）。这里统一解码为 3 通道 RGB，稳妥可靠。
 * 返回 {dict, data} 或 null。
 */
function pngToPdfImage(buf, w, h) {
  let off = 8, colorType = null, bitDepth = null, interlace = null;
  let idat = [], palette = null, trns = null;
  while (off + 12 <= buf.length) {
    const len = buf.readUInt32BE(off);
    const type = buf.toString('ascii', off + 4, off + 8);
    const dataOff = off + 8;
    if (type === 'IHDR') {
      w = buf.readUInt32BE(dataOff); h = buf.readUInt32BE(dataOff + 4);
      bitDepth = buf[dataOff + 8]; colorType = buf[dataOff + 9]; interlace = buf[dataOff + 12];
    } else if (type === 'PLTE') palette = buf.subarray(dataOff, dataOff + len);
    else if (type === 'tRNS') trns = buf.subarray(dataOff, dataOff + len);
    else if (type === 'IDAT') idat.push(buf.subarray(dataOff, dataOff + len));
    else if (type === 'IEND') break;
    off += 12 + len;
  }
  if (colorType == null || bitDepth !== 8 || !idat.length || interlace) return null; // 不支持隔行
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colorType];
  if (!channels) return null;

  let inflated;
  try { inflated = zlib.inflateSync(Buffer.concat(idat)); } catch (e) { return null; }
  const stride = w * channels;
  const expect = h * (stride + 1);
  if (inflated.length < expect) return null;

  // 逐行反滤波（PNG filter type 0..4），得到原始通道数据
  const raw = Buffer.allocUnsafe(h * stride);
  let p = 0;
  for (let y = 0; y < h; y++) {
    const ft = inflated[p++];
    const rowOff = y * stride;
    const prevOff = rowOff - stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= channels ? raw[rowOff + x - channels] : 0;
      const b = y > 0 ? raw[prevOff + x] : 0;
      const c = (y > 0 && x >= channels) ? raw[prevOff + x - channels] : 0;
      let v = inflated[p++];
      if (ft === 1) v = (v + a) & 255;
      else if (ft === 2) v = (v + b) & 255;
      else if (ft === 3) v = (v + ((a + b) >> 1)) & 255;
      else if (ft === 4) {
        const pp = a + b - c, pa = Math.abs(pp - a), pb = Math.abs(pp - b), pc = Math.abs(pp - c);
        const pr = (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
        v = (v + pr) & 255;
      } else if (ft !== 0) return null;
      raw[rowOff + x] = v;
    }
  }

  // 转成 3 通道 RGB，alpha 合成到白底
  const rgb = Buffer.allocUnsafe(w * h * 3);
  let si = 0, di = 0;
  for (let i = 0; i < w * h; i++) {
    let r, g, bl, al = 255;
    if (colorType === 0) { r = g = bl = raw[si++]; }
    else if (colorType === 2) { r = raw[si++]; g = raw[si++]; bl = raw[si++]; }
    else if (colorType === 3) {
      const idx = raw[si++]; const po = idx * 3;
      r = palette && po + 2 < palette.length ? palette[po] : 255;
      g = palette && po + 2 < palette.length ? palette[po + 1] : 255;
      bl = palette && po + 2 < palette.length ? palette[po + 2] : 255;
    } else if (colorType === 4) { r = g = bl = raw[si++]; al = raw[si++]; }
    else { r = raw[si++]; g = raw[si++]; bl = raw[si++]; al = raw[si++]; } // colorType 6 RGBA
    if (al !== 255) {
      r = Math.round(r * al / 255 + 255 * (1 - al / 255));
      g = Math.round(g * al / 255 + 255 * (1 - al / 255));
      bl = Math.round(bl * al / 255 + 255 * (1 - al / 255));
    }
    rgb[di++] = r; rgb[di++] = g; rgb[di++] = bl;
  }
  const def = zlib.deflateSync(rgb, { level: 6 });
  const dict = `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length ${def.length} >>`;
  return { dict: dict, data: def };
}

function generatePdf(text, title, style, opts) {
  ACTIVE = (style && style.ok !== false) ? style : null;
  opts = opts || {};
  // 从原简历提取的证件照（JPEG/PNG）：优先 JPEG，直接以 DCTDecode 嵌入
  var photo = null;
  if (Array.isArray(opts.images)) {
    var candidates = [];
    opts.images.forEach(function (im) {
      if (!im || !im.data || im.data.length < 200) return;
      var sz = null, isJpeg = /jpe?g/i.test(im.ext || im.name || ''), isPng = /png/i.test(im.ext || im.name || '');
      if (isJpeg) sz = jpegSize(im.data);
      else if (isPng && im.data.toString('latin1', 1, 4) === 'PNG') {
        try { sz = { w: im.data.readUInt32BE(16), h: im.data.readUInt32BE(20) }; } catch (e) {}
      }
      if (!sz || sz.w <= 20 || sz.h <= 20) return;
      // 证件照评分：接近正方形(宽高比~0.75-1.1)、尺寸中等(80-600px)得分高；
      // 超大背景图(>700px)或极端长条图得分极低。
      var ratio = sz.w / sz.h;
      var score = 0;
      if (ratio >= 0.7 && ratio <= 1.15) score += 100;       // 证件照多为近正方形
      else if (ratio >= 0.5 && ratio <= 1.5) score += 50;
      var maxSide = Math.max(sz.w, sz.h), minSide = Math.min(sz.w, sz.h);
      if (minSide >= 80 && maxSide <= 600) score += 60;      // 典型证件照尺寸
      else if (minSide >= 60 && maxSide <= 900) score += 20;
      if (maxSide > 700) score -= 80;                        // 整页/装饰大图，重罚
      if (minSide < 60) score -= 40;                         // 小图标，重罚
      if (isJpeg) score += 5;                                // JPEG 直接嵌入略优
      candidates.push({ im: im, w: sz.w, h: sz.h, png: !isJpeg, score: score });
    });
    candidates.sort(function (a, b) { return b.score - a.score; });
    if (candidates.length && candidates[0].score > 0) {
      var best = candidates[0];
      photo = { data: best.im.data, w: best.w, h: best.h, png: best.png };
    }
  }
  var font = ensureFontFor(text);
  // 字体子集化：只保留本次实际要渲染的字符（正文 + 标题 + 页脚），把字体从 MB 级降到几十 KB。
  // 这样 PDF 体积与生成耗时（deflate 大字体）都下降一个数量级，公网下载显著变快。
  var FOOTER = '本智能体仅作求职辅助，内容以个人思考为主；【】内请按真实经历填写';
  font = subsetFontForText(font, String(text || '') + String(title || '') + FOOTER);
  // 页面尺寸/边距：带原文件时用原值（pt），否则 A4 默认
  let PW = 595.28, PH = 841.89;
  let ML = 52, MR = 52, MT = 56, MB = 52;
  if (ACTIVE) {
    if (ACTIVE.pageW > 200 && ACTIVE.pageH > 200) { PW = ACTIVE.pageW; PH = ACTIVE.pageH; }
    // 原 docx 的边距 twips -> pt
    if (ACTIVE.marginTwips) {
      const mt = ACTIVE.marginTwips, tw = v => (v != null ? v / 20 : null);
      if (tw(mt.left) > 8) ML = Math.min(tw(mt.left), PW / 3);
      if (tw(mt.right) > 8) MR = Math.min(tw(mt.right), PW / 3);
      if (tw(mt.top) > 8) MT = Math.min(tw(mt.top), PH / 3);
      if (tw(mt.bottom) > 8) MB = Math.min(tw(mt.bottom), PH / 3);
    }
  }
  const usable = PW - ML - MR;

  const rawLines = mdToLines(text);
  // 展开为带样式的物理行
  const flowed = [];
  rawLines.forEach(l => {
    if (l.rule) { flowed.push({ rule: true }); return; }
    if (!l.text) { flowed.push({ blank: true, gap: l.gap || 6 }); return; }
    const size = l.size || 10.5;
    const wrapped = wrapText(l.text, font, usable / size);
    wrapped.forEach((lineText, i) => {
      flowed.push({
        text: lineText, size,
        bold: !!l.bold, color: l.color,
        indent: (i === 0 ? (l.indent || 0) : (l.indent || 0) + 10),
        before: i === 0 ? (l.before || 0) : 2,
        after: i === wrapped.length - 1 ? (l.after || 0) : 0
      });
    });
  });

  // 带原文件样式时：不额外绘制 AI 模板大标题（原样式由正文自身表达）
  const noTemplateTitle = !!(ACTIVE && opts && opts.customTitle);

  // 证件照尺寸与位置（右上角），正文首段避开它
  let photoBox = null;
  if (photo) {
    const fitW = 92;                                   // 目标宽度 pt（约 3.2cm）
    const imgW = fitW, imgH = Math.round(fitW * photo.h / photo.w);
    const px = PW - MR - imgW - 4, py = PH - MT - imgH;
    photoBox = { x: px, y: py, w: imgW, h: imgH };
  }

  // 分页布局
  const pages = [];
  let cur = [], y = PH - MT;
  let firstBody = true; // 首行正文标记：有证件照时把首行下移到照片下方
  if (!noTemplateTitle) {
    cur.push({ op: 'title', text: stripInline(title || 'AI 润色简历'), size: 18 });
    y -= 30;
  }
  flowed.forEach(item => {
    if (item.rule) {
      y -= 7;
      if (y < MB) { pages.push(cur); cur = []; y = PH - MT; }
      cur.push({ op: 'rule', y }); y -= 7; return;
    }
    if (item.blank) { y -= (item.gap || 6); return; }
    const leading = item.size * 1.55;
    y -= (item.before || 0);
    // 首页首个正文行：若右上角有证件照且该行会与照片水平重叠，则下移到照片下方
    if (photoBox && firstBody && pages.length === 0) {
      if (y > photoBox.y - leading) y = photoBox.y - 12;
    }
    if (y - leading < MB) { pages.push(cur); cur = []; y = PH - MT; firstBody = false; }
    cur.push({ op: 'text', text: item.text, x: ML + (item.indent || 0), y, size: item.size, bold: item.bold, color: item.color });
    y -= leading + (item.after || 0);
    firstBody = false;
  });
  pages.push(cur);

  // 内容流
  const streams = pages.map((items, pi) => {
    let s = '';
    // 首页绘制证件照（右上角）
    if (photoBox && pi === 0) {
      s += `q\n${photoBox.w} 0 0 ${photoBox.h} ${photoBox.x.toFixed(1)} ${photoBox.y.toFixed(1)} cm\n/Im0 Do\nQ\n`;
    }
    items.forEach(it => {
      if (it.op === 'title') return; // 标题单独在每页顶部绘制见下
      if (it.op === 'rule') {
        s += `0.85 w 0.78 0.82 0.87 RG ${ML} ${it.y} m ${PW - MR} ${it.y} l S\n`;
        return;
      }
      if (it.op === 'text') {
        const headRgb = hexRg(styleHeadingColor(), '1F4E9C');
        const bodyRgb = (ACTIVE && ACTIVE.bodyColor && /^[0-9A-Fa-f]{6}$/.test(ACTIVE.bodyColor))
          ? hexRg(ACTIVE.bodyColor, '') : null;
        const rgb = it.color === styleHeadingColor() ? (headRgb || '0.12 0.31 0.61 rg')
          : (bodyRgb || (it.bold ? '0.10 0.12 0.16 rg' : '0.16 0.18 0.22 rg'));
        s += `${rgb}\nBT /F1 ${it.size.toFixed(1)} Tf 1 0 0 1 ${it.x.toFixed(1)} ${it.y.toFixed(1)} Tm <${textToHex(it.text, font)}> Tj ET\n`;
      }
    });
    if (!noTemplateTitle) {
      // 每页页眉标题
      s += `${hexRg(styleHeadingColor(), '1F4E9C')}\nBT /F1 15 Tf 1 0 0 1 ${ML} ${PH - 32} Tm <${textToHex(stripInline(title || 'AI 润色简历'), font)}> Tj ET\n`;
      s += `0.60 0.62 0.66 rg\nBT /F1 7.5 Tf 1 0 0 1 ${ML} 28 Tm <${textToHex('本智能体仅作求职辅助，内容以个人思考为主；【】内请按真实经历填写', font)}> Tj ET\n`;
    }
    return s;
  });

  // /W 宽度数组：按连续 gid 分组压缩
  const wPairs = [];
  let start = 0, prev = font.widths[0] || 500;
  for (let gid = 1; gid <= font.cmap.maxGid; gid++) {
    const w = font.widths[gid] || 500;
    if (w !== prev) { wPairs.push(`${start} [${prev}]`); start = gid; prev = w; }
  }
  wPairs.push(`${start} [${prev}]`);
  const W = '[ ' + wPairs.join(' ') + ' ]';

  // ToUnicode CMap：只映射本次实际用到的字符，避免子集字体产生庞大的映射表
  const usedMap = new Map(); // gid -> unicode code
  const allText = String(text || '') + '\n' + String(title || '') + '\n本智能体仅作求职辅助，内容以个人思考为主；【】内请按真实经历填写';
  for (const ch of allText) {
    const code = ch.codePointAt(0);
    if (code < 0x20) continue;
    const gid = font.cmap.map[code];
    if (gid && !usedMap.has(gid)) usedMap.set(gid, code);
  }
  const used = [];
  usedMap.forEach((code, gid) => used.push(`<${hex4(gid)}> <${hex4(code)}>`));
  const cmapStream = [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin', 'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def', '/CMapType 2 def',
    '1 begincodespacerange', '<0000> <FFFF>', 'endcodespacerange',
    `${used.length} beginbfchar`, ...used, 'endbfchar',
    'endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end'
  ].join('\n');

  const asc = font.ascender || 880, desc = font.descender || -120;
  const bbox = font.fontBBox.map(v => Math.round(v * 1000 / font.unitsPerEm));

  // 对象编号规划（先定编号，再填内容）
  // 1=Catalog 2=Pages 3..(2+N)=Page 其后 N 个=Contents 再后=Font/CIDFont/ToUnicode/FontDescriptor/FontFile2/Info
  const N = streams.length;
  const NUM = {
    catalog: 1,
    pages: 2,
    pageStart: 3,                 // 连续 N 个
    contentStart: 3 + N,          // 连续 N 个
    font: 3 + 2 * N,
    cidFont: 4 + 2 * N,
    toUnicode: 5 + 2 * N,
    fontDesc: 6 + 2 * N,
    fontFile: 7 + 2 * N,
    info: 8 + 2 * N,
    imgData: 9 + 2 * N            // 证件照图像流（仅首页）：JPEG=DCTDecode / PNG=FlateDecode+预测器
  };
  // 证件照图像对象：先真正生成成功，才把它计入对象表并在页面里引用。
  // 绝不能出现"页面引用 /Im0 但对象不存在" —— 那会写出 "undefined 0 obj" 损坏整个 PDF。
  const binaryObjs = {};
  if (photo && !photo.png) {
    binaryObjs[NUM.imgData] = {
      dict: `<< /Type /XObject /Subtype /Image /Width ${photo.w} /Height ${photo.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${photo.data.length} >>`,
      data: photo.data
    };
  } else if (photo && photo.png) {
    try {
      const raw = pngToPdfImage(photo.data, photo.w, photo.h);
      if (raw) binaryObjs[NUM.imgData] = raw;
    } catch (e) {}
  }
  const hasPhoto = !!binaryObjs[NUM.imgData];
  // 图像绘制指令只有在图像对象确实存在时才有效，否则清理掉避免悬空引用
  if (!hasPhoto) streams[0] = streams[0].replace(/q\n[\d.]+ 0 0 [\d.]+ [\d.]+ [\d.]+ cm\n\/Im0 Do\nQ\n/g, '');

  const TOTAL = hasPhoto ? NUM.imgData : NUM.info;

  const objs = {}; // num -> string body（FontFile2 的为 null，单独处理）
  objs[NUM.catalog] = `<< /Type /Catalog /Pages ${NUM.pages} 0 R >>`;
  const kids = Array.from({ length: N }, (_, i) => `${NUM.pageStart + i} 0 R`).join(' ');
  objs[NUM.pages] = `<< /Type /Pages /Kids [${kids}] /Count ${N} >>`;

  streams.forEach((stream, i) => {
    const xobj = (hasPhoto && i === 0) ? ' /XObject << /Im0 ' + NUM.imgData + ' 0 R >>' : '';
    objs[NUM.pageStart + i] = `<< /Type /Page /Parent ${NUM.pages} 0 R /MediaBox [0 0 ${PW.toFixed(2)} ${PH.toFixed(2)}] /Resources << /Font << /F1 ${NUM.font} 0 R >>${xobj} >> /Contents ${NUM.contentStart + i} 0 R >>`;
    objs[NUM.contentStart + i] = `<< /Length ${Buffer.byteLength(stream, 'latin1')} >>\nstream\n${stream}endstream`;
  });

  objs[NUM.font] = `<< /Type /Font /Subtype /Type0 /BaseFont /SimHei /Encoding /Identity-H /DescendantFonts [${NUM.cidFont} 0 R] /ToUnicode ${NUM.toUnicode} 0 R >>`;
  objs[NUM.cidFont] = `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /SimHei /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${NUM.fontDesc} 0 R /DW 1000 /W ${W} >>`;
  objs[NUM.toUnicode] = `<< /Length ${Buffer.byteLength(cmapStream, 'latin1')} >>\nstream\n${cmapStream}\nendstream`;
  objs[NUM.fontDesc] = `<< /Type /FontDescriptor /FontName /SimHei /Flags 6 /FontBBox [${bbox.join(' ')}] /ItalicAngle 0 /Ascent ${asc} /Descent ${desc} /CapHeight ${asc} /StemV 80 /FontFile2 ${NUM.fontFile} 0 R >>`;
  objs[NUM.info] = `<< /Title (${pdfEscape(title || 'AI 润色简历')}) /Producer (AI 求职小助手) >>`;
  // FontFile2 二进制流：用 FlateDecode 压缩子集字体，PDF 体积再降约一半
  const fontBytes = font.buf.subarray(font.base);
  const fontStream = zlib.deflateSync(fontBytes, { level: 9 });
  objs[NUM.fontFile] = null;

  // 序列化：按编号从小到大输出，记录每个对象偏移
  const chunks = [];
  const offsets = new Array(TOTAL + 1).fill(0);
  const header = Buffer.from('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n', 'latin1');
  chunks.push(header);
  let running = header.length;
  for (let num = 1; num <= TOTAL; num++) {
    offsets[num] = running;
    if (num === NUM.fontFile) {
      const pfx = Buffer.from(`${num} 0 obj\n<< /Length ${fontStream.length} /Length1 ${fontBytes.length} /Filter /FlateDecode >>\nstream\n`, 'latin1');
      const sfx = Buffer.from(`\nendstream\nendobj\n`, 'latin1');
      chunks.push(pfx, fontStream, sfx);
      running += pfx.length + fontStream.length + sfx.length;
    } else if (binaryObjs[num]) {
      const bo = binaryObjs[num];
      const pfx = Buffer.from(`${num} 0 obj\n${bo.dict}\nstream\n`, 'latin1');
      const sfx = Buffer.from(`\nendstream\nendobj\n`, 'latin1');
      chunks.push(pfx, bo.data, sfx);
      running += pfx.length + bo.data.length + sfx.length;
    } else {
      const b = Buffer.from(`${num} 0 obj\n${objs[num]}\nendobj\n`, 'latin1');
      chunks.push(b);
      running += b.length;
    }
  }
  const bodyBuf = Buffer.concat(chunks);
  const xrefOffset = bodyBuf.length;
  const xrefLines = ['xref', `0 ${TOTAL + 1}`, '0000000000 65535 f '];
  for (let i = 1; i <= TOTAL; i++) xrefLines.push(`${String(offsets[i]).padStart(10, '0')} 00000 n `);
  const trailer = `trailer\n<< /Size ${TOTAL + 1} /Root 1 0 R /Info ${NUM.info} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;

  return Buffer.concat([bodyBuf, Buffer.from(xrefLines.join('\n') + '\n', 'latin1'), Buffer.from(trailer, 'latin1')]);
}

module.exports = { generatePdf };
