/**
 * 零依赖 .docx 生成器：把（轻量 Markdown）文本导出为可被 Word / WPS 打开的标准 docx。
 * 仅用 Node 内置能力，手写 ZIP(store) + OOXML，不引入任何第三方包。
 */

// ---------- CRC32 ----------
var CRC_TABLE = (function () {
  var t = new Uint32Array(256);
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  var c = 0xFFFFFFFF;
  for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

// ---------- 极简 ZIP (store, 不压缩) ----------
function buildZip(files) {
  // files: [{name, data:Buffer}]
  var chunks = [];
  var central = [];
  var offset = 0;

  files.forEach(function (f) {
    var nameBuf = Buffer.from(f.name, 'utf8');
    var data = f.data;
    var crc = crc32(data);

    var local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);   // local file header signature
    local.writeUInt16LE(20, 4);           // version needed
    local.writeUInt16LE(0x0800, 6);       // flags: UTF-8 filename
    local.writeUInt16LE(0, 8);            // compression: store
    local.writeUInt16LE(0, 10);           // mod time
    local.writeUInt16LE(0, 12);           // mod date
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);           // extra length
    chunks.push(local, nameBuf, data);

    var cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);      // central dir signature
    cd.writeUInt16LE(20, 4);              // version made by
    cd.writeUInt16LE(20, 6);              // version needed
    cd.writeUInt16LE(0x0800, 8);          // flags
    cd.writeUInt16LE(0, 10);              // compression
    cd.writeUInt16LE(0, 12);              // mod time
    cd.writeUInt16LE(0, 14);              // mod date
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(data.length, 20);
    cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt16LE(0, 30);              // extra
    cd.writeUInt16LE(0, 32);              // comment
    cd.writeUInt16LE(0, 34);              // disk number
    cd.writeUInt16LE(0, 36);              // internal attrs
    cd.writeUInt32LE(0, 38);              // external attrs
    cd.writeUInt32LE(offset, 42);         // local header offset
    central.push(cd, nameBuf);

    offset += local.length + nameBuf.length + data.length;
  });

  var centralSize = central.reduce(function (s, b) { return s + b.length; }, 0);
  var centralOffset = offset;
  var eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(centralOffset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat(chunks.concat(central, [eocd]));
}

// ---------- XML 转义 ----------
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// 当前导出样式（由 generateDocx 注入）：默认沿用原生蓝色模板；带 source 时套用原文件样式
var ACTIVE_STYLE = null;
function fontTag() {
  return ACTIVE_STYLE && ACTIVE_STYLE.font
    ? '<w:rFonts w:ascii="' + esc(ACTIVE_STYLE.font) + '" w:hAnsi="' + esc(ACTIVE_STYLE.font) + '" w:eastAsia="' + esc(ACTIVE_STYLE.font) + '"/>'
    : '<w:rFonts w:eastAsia="Microsoft YaHei"/>';
}
function bodyColor() { return (ACTIVE_STYLE && ACTIVE_STYLE.bodyColor) || '1F4E9C'; }
function headingColor() { return (ACTIVE_STYLE && ACTIVE_STYLE.headingColor) || '1F4E9C'; }
function bodySz() { return ACTIVE_STYLE && ACTIVE_STYLE.fontSizePt ? Math.round(ACTIVE_STYLE.fontSizePt * 2) : 22; }

// 从图片列表里挑出最像证件照的一张：接近正方形、尺寸中等；排除装饰大图与小图标
function pickPhoto(images) {
  if (!Array.isArray(images)) return null;
  var best = null, bestScore = -1e9;
  images.forEach(function (im) {
    if (!im || !im.data || im.data.length < 200) return;
    var w = 0, h = 0, isJpeg = /jpe?g/i.test(im.ext || im.name || ''), isPng = /png/i.test(im.ext || im.name || '');
    if (isPng && im.data.toString('latin1', 1, 4) === 'PNG') {
      try { w = im.data.readUInt32BE(16); h = im.data.readUInt32BE(20); } catch (e) {}
    } else if (isJpeg) {
      try {
        var b = im.data, i = 2;
        while (i < b.length - 9) {
          if (b[i] !== 0xFF) { i++; continue; }
          var m = b[i + 1];
          if (m === 0xD8 || m === 0xD9 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }
          var len = b.readUInt16BE(i + 2);
          if ((m >= 0xC0 && m <= 0xC3) || (m >= 0xC5 && m <= 0xC7) || (m >= 0xC9 && m <= 0xCB) || (m >= 0xCD && m <= 0xCF)) {
            h = b.readUInt16BE(i + 5); w = b.readUInt16BE(i + 7); break;
          }
          i += 2 + len;
        }
      } catch (e) {}
    }
    if (!w || !h || w <= 20 || h <= 20) return;
    var ratio = w / h, score = 0;
    if (ratio >= 0.7 && ratio <= 1.15) score += 100;
    else if (ratio >= 0.5 && ratio <= 1.5) score += 50;
    var mx = Math.max(w, h), mn = Math.min(w, h);
    if (mn >= 80 && mx <= 600) score += 60;
    else if (mn >= 60 && mx <= 900) score += 20;
    if (mx > 700) score -= 80;
    if (mn < 60) score -= 40;
    if (score > bestScore) {
      bestScore = score;
      best = { data: im.data, w: w, h: h, ext: isJpeg ? 'jpeg' : 'png' };
    }
  });
  return (best && bestScore > 0) ? best : null;
}

// 证件照 inline drawing（Word 原生支持 PNG/JPEG，无需重编码）
function photoXml(photo) {
  var maxW = 1100000; // EMU，约 1.15 英寸
  var cx = 9525 * photo.w, cy = 9525 * photo.h;
  if (cx > maxW) { cy = Math.round(cy * maxW / cx); cx = maxW; }
  return '<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:after="120"/></w:pPr><w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="1" name="Photo"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="photo"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdPhoto"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
}

// 处理行内 **加粗** 与 `代码`
function inlineRuns(text) {
  var out = '';
  var re = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  var last = 0, m;
  function run(txt, bold, code) {
    if (!txt) return;
    var rPr = fontTag();
    if (bold) rPr += '<w:b/>';
    if (code) rPr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas"/><w:shd w:fill="F2F2F2"/>';
    out += '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(txt) + '</w:t></w:r>';
  }
  while ((m = re.exec(text)) !== null) {
    run(text.slice(last, m.index), false, false);
    var tok = m[0];
    if (tok.charAt(0) === '*') run(tok.slice(2, -2), true, false);
    else run(tok.slice(1, -1), false, true);
    last = m.index + tok.length;
  }
  run(text.slice(last), false, false);
  if (!out) out = '<w:r><w:t xml:space="preserve"> </w:t></w:r>';
  return out;
}

function paragraph(text, opts) {
  opts = opts || {};
  var pPr = '<w:pPr>';
  if (opts.heading) {
    var sz = opts.heading === 1 ? 32 : (opts.heading === 2 ? 28 : 24);
    pPr += '<w:pStyle w:val="Heading' + Math.min(opts.heading, 3) + '"/>';
    pPr += '<w:spacing w:before="240" w:after="120"/>';
    pPr += '<w:rPr><w:b/><w:color w:val="' + headingColor() + '"/><w:sz w:val="' + sz + '"/></w:rPr>';
  } else if (opts.list) {
    pPr += '<w:ind w:left="' + (opts.listIndent || 360) + '"/>';
    pPr += '<w:spacing w:after="80"/>';
  } else {
    pPr += '<w:spacing w:after="120" w:line="360" w:lineRule="auto"/>';
  }
  pPr += '</w:pPr>';
  var runs = opts.heading
    ? '<w:r><w:rPr><w:b/><w:color w:val="' + headingColor() + '"/>' + fontTag() + '</w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>'
    : inlineRuns(text);
  var prefix = opts.bullet
    ? '<w:r><w:rPr><w:color w:val="' + headingColor() + '"/></w:rPr><w:t xml:space="preserve">' + opts.bullet + ' </w:t></w:r>'
    : '';
  return '<w:p>' + pPr + prefix + runs + '</w:p>';
}

// 把轻量 Markdown 文本转换为 document.xml 的 body 内容
function markdownToBody(src) {
  var lines = String(src == null ? '' : src).replace(/\r/g, '').split('\n');
  var body = '';
  lines.forEach(function (line) {
    var t = line.replace(/\s+$/, '');
    if (!t.trim()) { body += '<w:p><w:spacing w:after="60"/></w:p>'; return; }
    if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(t)) {
      body += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="D0D7DE"/></w:pBdr><w:spacing w:before="120" w:after="120"/></w:pPr></w:p>';
      return;
    }
    var h = t.match(/^\s*(#{1,4})\s+(.*)$/);
    if (h) { body += paragraph(h[2], { heading: Math.min(h[1].length, 3) }); return; }
    var ol = t.match(/^\s*(\d+)[.)]\s+(.*)$/);
    if (ol) { body += paragraph(ol[2], { list: true, bullet: ol[1] + '.' }); return; }
    var ul = t.match(/^\s*[-*·]\s+(.*)$/);
    if (ul) { body += paragraph(ul[1], { list: true, bullet: '•' }); return; }
    body += paragraph(t);
  });
  return body;
}

function buildDocumentXml(bodyInner, title, opts) {
  opts = opts || {};
  // 页边距 / 纸张：带原文件时用原值（twips），否则 A4 默认
  var mt = 1417, mr = 1417, mb = 1417, ml = 1417, pw = 11906, ph = 16838;
  var st = ACTIVE_STYLE;
  if (st && st.marginTwips) {
    if (st.marginTwips.top != null) mt = st.marginTwips.top;
    if (st.marginTwips.right != null) mr = st.marginTwips.right;
    if (st.marginTwips.bottom != null) mb = st.marginTwips.bottom;
    if (st.marginTwips.left != null) ml = st.marginTwips.left;
  }
  if (st && st.pageTwips) {
    if (st.pageTwips.w != null) pw = st.pageTwips.w;
    if (st.pageTwips.h != null) ph = st.pageTwips.h;
  }
  var titleSize = 36;
  // customTitle：润色正文自带 # 标题，不再单独输出标题段，避免"AI 润色简历/个人简历"双标题
  var titlePara = '';
  if (!opts.customTitle) {
    titlePara = '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>' +
      '<w:r><w:rPr><w:b/><w:color w:val="' + headingColor() + '"/><w:sz w:val="' + titleSize + '"/>' + fontTag() + '</w:rPr>' +
      '<w:t xml:space="preserve">' + esc(title || 'AI 润色简历') + '</w:t></w:r></w:p>';
  }
  // 证件照放在文档最前
  var photoPara = opts.photo ? photoXml(opts.photo) : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:body>' + titlePara + photoPara + bodyInner +
    '<w:sectPr><w:pgSz w:w="' + pw + '" w:h="' + ph + '"/>' +
    '<w:pgMar w:top="' + mt + '" w:right="' + mr + '" w:bottom="' + mb + '" w:left="' + ml + '" w:header="720" w:footer="720" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>';
}

var CONTENT_TYPES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
  '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
  '</Types>';

var RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
  '</Relationships>';

var DOC_RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '</Relationships>';

var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
  '</w:styles>';

function coreXml(title) {
  var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:title>' + esc(title || 'AI 润色简历') + '</dc:title>' +
    '<dc:creator>AI 求职小助手</dc:creator>' +
    '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>' +
    '</cp:coreProperties>';
}
var APP_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
  '<Application>AI 求职小助手</Application></Properties>';

/**
 * 生成 docx Buffer
 * @param {string} text 轻量 Markdown 文本
 * @param {string} title 文档标题
 * @returns {Buffer}
 */
function stylesXml() {
  var s = STYLES;
  var st = ACTIVE_STYLE;
  if (st && (st.font || st.fontSizePt)) {
    var font = st.font || 'Microsoft YaHei';
    var sz = st.fontSizePt ? Math.round(st.fontSizePt * 2) : 22;
    s = s.replace('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>',
      '<w:rFonts w:ascii="' + esc(font) + '" w:hAnsi="' + esc(font) + '" w:eastAsia="' + esc(font) + '"/>');
    s = s.replace('<w:sz w:val="22"/><w:szCs w:val="22"/>',
      '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>');
  }
  return s;
}

/**
 * 生成 docx Buffer
 * @param {string} text 轻量 Markdown 文本（润色后简历正文）
 * @param {string} title 文档标题
 * @param {object} style 可选：从原简历提取的样式 {font,fontSizePt,bodyColor,headingColor,marginTwips,pageTwips}
 * @param {object} opts 可选：{customTitle:true} 时用较小标题字号
 */
function generateDocx(text, title, style, opts) {
  opts = opts || {};
  ACTIVE_STYLE = (style && style.ok !== false && style.font) ? style : null;
  var body = markdownToBody(text);
  var docTitle = (ACTIVE_STYLE && opts.customTitle) ? '' : (title || 'AI 润色简历');
  // 证件照：从原文件图片中挑选（接近正方形、中等尺寸），嵌入新文档
  var photo = pickPhoto(opts.images);
  var runOpts = { customTitle: opts.customTitle, photo: photo };
  var contentTypes = CONTENT_TYPES;
  var docRels = DOC_RELS;
  var files = [
    { name: '[Content_Types].xml', data: Buffer.from(contentTypes, 'utf8') },
    { name: '_rels/.rels', data: Buffer.from(RELS, 'utf8') },
    { name: 'word/document.xml', data: Buffer.from(buildDocumentXml(body, docTitle, runOpts), 'utf8') },
    { name: 'word/styles.xml', data: Buffer.from(stylesXml(), 'utf8') },
    { name: 'docProps/core.xml', data: Buffer.from(coreXml(title), 'utf8') },
    { name: 'docProps/app.xml', data: Buffer.from(APP_XML, 'utf8') }
  ];
  if (photo) {
    var ext = photo.ext === 'jpeg' ? 'jpeg' : 'png';
    var ct = ext === 'jpeg' ? 'image/jpeg' : 'image/png';
    contentTypes = CONTENT_TYPES.replace('</Types>', '<Default Extension="' + ext + '" ContentType="' + ct + '"/></Types>');
    docRels = DOC_RELS.replace('</Relationships>',
      '<Relationship Id="rIdPhoto" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo.' + ext + '"/></Relationships>');
    files[0] = { name: '[Content_Types].xml', data: Buffer.from(contentTypes, 'utf8') };
    files.push({ name: 'word/media/photo.' + ext, data: photo.data });
  }
  files.splice(3, 0, { name: 'word/_rels/document.xml.rels', data: Buffer.from(docRels, 'utf8') });
  var out = buildZip(files);
  ACTIVE_STYLE = null;
  return out;
}

module.exports = { generateDocx: generateDocx };
