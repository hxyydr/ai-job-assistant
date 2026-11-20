/**
 * 零依赖：从用户上传的原始简历文件（docx / pdf）中提取样式参数，
 * 供导出器套用，使「润色后下载的文件」字体/字号/颜色/页边距尽量与原文件一致。
 * docx 是 zip：用 Node 内置 zlib 解 central directory，读 word/document.xml 与 word/styles.xml。
 * pdf：从原始字节里正则提取 /MediaBox、/BaseFont、Tf 字号等——PDF 样式提取较粗糙，
 *      能拿到页面尺寸与正文字号即可；字体受限于本应用已嵌入的字体，做名称匹配。
 */
'use strict';
const zlib = require('zlib');

// ---------- 极简 ZIP 读取 ----------
function unzip(buffer) {
  const files = {};
  // 找 EOCD（从尾部向前）
  let eocd = -1;
  for (let i = buffer.length - 22; i >= 0; i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return files;
  const n = buffer.readUInt16LE(eocd + 10);
  let off = buffer.readUInt32LE(eocd + 16);
  for (let i = 0; i < n; i++) {
    if (buffer.readUInt32LE(off) !== 0x02014b50) break;
    const comp = buffer.readUInt16LE(off + 10);
    const csize = buffer.readUInt32LE(off + 20);
    const nl = buffer.readUInt16LE(off + 28);
    const el = buffer.readUInt16LE(off + 30);
    const cl = buffer.readUInt16LE(off + 32);
    const lho = buffer.readUInt32LE(off + 42);
    const name = buffer.toString('utf8', off + 46, off + 46 + nl);
    const lnl = buffer.readUInt16LE(lho + 26);
    const lel = buffer.readUInt16LE(lho + 28);
    const dstart = lho + 30 + lnl + lel;
    const raw = buffer.subarray(dstart, dstart + csize);
    try { files[name] = comp === 8 ? zlib.inflateRawSync(raw) : raw; } catch (e) { files[name] = raw; }
    off += 46 + nl + el + cl;
  }
  return files;
}

function pick(xml, re, def) { const m = xml.match(re); return m ? m[1] : def; }

/**
 * 从 docx 提取样式：正文默认字体(eastAsia)、字号(w:sz 半磅)、颜色、标题色、页边距(twips/转 pt)。
 * 返回 {font, fontSizePt, bodyColor, headingColor, marginTwips:{top,right,bottom,left}, ok:true}
 */
function extractDocxStyle(buffer) {
  try {
    const files = unzip(buffer);
    const docXml = files['word/document.xml'] ? files['word/document.xml'].toString('utf8') : '';
    const styXml = files['word/styles.xml'] ? files['word/styles.xml'].toString('utf8') : '';
    if (!docXml) return { ok: false };
    const all = styXml + '\n' + docXml;
    // 正文字体：优先 styles.xml 的 docDefaults eastAsia，其次任意 rFonts eastAsia
    let font = pick(all, /w:rFonts[^>]*w:eastAsia="([^"]+)"/, '') || pick(all, /w:rFonts[^>]*w:ascii="([^"]+)"/, '');
    // 正文字号：w:sz 是半磅
    let szHalf = parseInt(pick(all, /<w:sz w:val="(\d+)"/, ''), 10) || 21; // 默认 10.5pt
    // 正文颜色、标题颜色
    let bodyColor = pick(all, /<w:color w:val="([0-9A-Fa-f]{6})"/, '');
    let headingColor = pick(styXml, /Heading1[\s\S]{0,400}?<w:color w:val="([0-9A-Fa-f]{6})"/, '') || bodyColor;
    // 页边距 twips
    let pgMar = docXml.match(/<w:pgMar\b([^>]*)\/>/);
    let mt = pgMar ? parseInt(pick(pgMar[1], /w:top="(\d+)"/, ''), 10) : null;
    let mr = pgMar ? parseInt(pick(pgMar[1], /w:right="(\d+)"/, ''), 10) : null;
    let mb = pgMar ? parseInt(pick(pgMar[1], /w:bottom="(\d+)"/, ''), 10) : null;
    let ml = pgMar ? parseInt(pick(pgMar[1], /w:left="(\d+)"/, ''), 10) : null;
    // 纸张 twips -> PDF pt（1 twip = 1/20 pt）
    let pgSz = docXml.match(/<w:pgSz\b([^>]*)\/>/);
    let pw = pgSz ? parseInt(pick(pgSz[1], /w:w="(\d+)"/, ''), 10) : null;
    let ph = pgSz ? parseInt(pick(pgSz[1], /w:h="(\d+)"/, ''), 10) : null;
    return {
      ok: true,
      font: font || '',
      fontSizePt: szHalf / 2,
      bodyColor: bodyColor || '',
      headingColor: headingColor || '',
      marginTwips: { top: mt, right: mr, bottom: mb, left: ml },
      pageTwips: { w: pw, h: ph }
    };
  } catch (e) { return { ok: false, error: String(e && e.message || e) }; }
}

// 已知字体名 -> 候选 ttf/ttc 文件（Windows 系统字体）
const FONT_FILES = {
  '黑体': 'C:\\Windows\\Fonts\\simhei.ttf', 'SimHei': 'C:\\Windows\\Fonts\\simhei.ttf',
  '宋体': 'C:\\Windows\\Fonts\\simsun.ttc', 'SimSun': 'C:\\Windows\\Fonts\\simsun.ttc',
  '微软雅黑': 'C:\\Windows\\Fonts\\msyh.ttc', 'Microsoft YaHei': 'C:\\Windows\\Fonts\\msyh.ttc',
  'Microsoft YaHei UI': 'C:\\Windows\\Fonts\\msyh.ttc',
  '等线': 'C:\\Windows\\Fonts\\msyh.ttc', 'DengXian': 'C:\\Windows\\Fonts\\msyh.ttc'
};

/** 把字体中文/英文名解析成实际字体文件路径，找不到返回 null（导出器用默认字体） */
function resolveFontPath(name) {
  if (!name) return null;
  const key = String(name).trim();
  if (FONT_FILES[key]) return FONT_FILES[key];
  for (const k of Object.keys(FONT_FILES)) {
    if (key.indexOf(k) >= 0) return FONT_FILES[k];
  }
  return null;
}

/**
 * 从 pdf 提取样式：页面尺寸（MediaBox）、正文字号（Tf）、颜色。
 * 返回 {pageW,pageH,fontSizePt,bodyColor,family,ok:true}
 */
function extractPdfStyle(buffer) {
  try {
    const s = buffer.toString('latin1');
    // MediaBox：可能出现多次，取第一个
    let pageW = null, pageH = null;
    const mb = s.match(/\/MediaBox\s*\[\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)/);
    if (mb) { pageW = parseFloat(mb[3]) - parseFloat(mb[1]); pageH = parseFloat(mb[4]) - parseFloat(mb[2]); }
    // 正文字号：统计所有 "N Tf" 的字号，取出现最多的（0 号字体命令常为标题，过滤掉偏大的）
    const sizes = {};
    const tfRe = /([\d.]+)\s+Tf/g; let m;
    while ((m = tfRe.exec(s)) !== null) { const v = Math.round(parseFloat(m[1]) * 10) / 10; if (v >= 6 && v <= 24) sizes[v] = (sizes[v] || 0) + 1; }
    let fontSizePt = null, best = 0;
    Object.keys(sizes).forEach(k => { if (sizes[k] > best) { best = sizes[k]; fontSizePt = parseFloat(k); } });
    // 字体名
    let family = null;
    const bf = s.match(/\/BaseFont\s*\/([A-Za-z0-9+\-,]+)/);
    if (bf) {
      const raw = bf[1].replace(/^[A-Z]{6}\+/, '');
      if (/SimSun|Song|宋/i.test(raw)) family = '宋体';
      else if (/YaHei|微软雅黑/i.test(raw)) family = '微软雅黑';
      else if (/Hei|黑/i.test(raw)) family = '黑体';
    }
    return { ok: true, pageW, pageH, fontSizePt, family, bodyColor: '' };
  } catch (e) { return { ok: false, error: String(e && e.message || e) }; }
}

/**
 * 从原始 PDF 提取内嵌图片（供导出时保留证件照）。
 * 支持 DCTDecode(JPEG) 与 FlateDecode(需为 image XObject)。返回 [{name,data,ext}]。
 */
function extractPdfImages(buffer) {
  const out = [];
  try {
    const b = buffer;
    // 逐个查找 image XObject：/Subtype /Image ... stream ... endstream
    const s = b.toString('latin1');
    const re = /\/Subtype\s*\/Image\b([\s\S]{0,600}?)stream\r?\n/g;
    let m;
    while ((m = re.exec(s)) !== null && out.length < 8) {
      const head = m[1];
      const dataStart = m.index + m[0].length;
      const end = s.indexOf('endstream', dataStart);
      if (end < 0) continue;
      let data = b.subarray(dataStart, end);
      // 去掉 endstream 前一个换行
      if (data.length && data[data.length - 1] === 0x0A) data = data.subarray(0, data.length - 1);
      if (data.length && data[data.length - 1] === 0x0D) data = data.subarray(0, data.length - 1);
      let ext = null;
      if (/\/DCTDecode/.test(head)) ext = 'jpeg';
      else if (/\/FlateDecode/.test(head)) ext = 'png-raw'; // PDF 原生 zlib 流，非 PNG 文件
      if (ext === 'jpeg' && data[0] === 0xFF && data[1] === 0xD8) out.push({ name: 'pdf-photo-' + out.length, data: Buffer.from(data), ext: 'jpeg' });
      // png-raw 无法直接当 PNG 用，跳过（docx 证件照路径已覆盖主要场景）
    }
  } catch (e) {}
  return out;
}

/** 统一入口：根据扩展名提取样式 */
function extractStyle(buffer, ext) {
  if (ext === 'docx') return extractDocxStyle(buffer);
  if (ext === 'pdf') return extractPdfStyle(buffer);
  return { ok: false };
}

module.exports = { extractStyle, extractDocxStyle, extractPdfStyle, resolveFontPath, unzip, extractPdfImages };
