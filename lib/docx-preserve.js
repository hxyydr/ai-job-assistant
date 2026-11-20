/**
 * 保真导出：基于用户上传的原始 docx，只替换正文文字，完整保留原文件的
 * 样式（styles.xml）、复选框编号（numbering.xml + w:numPr）、字体字号颜色
 * （w:rPr）、段落排版（w:pPr）、以及所有图片（word/media/*，含证件照）。
 *
 * 做法：
 *   1. 解压原 docx（复用 source-style 的 unzip）
 *   2. 解析 document.xml 的段落序列，提取每段纯文字
 *   3. 把润色后的正文按行拆分，与原文段落做「顺序 + 相似度」对齐
 *   4. 对匹配上的段落，仅替换其 <w:t> 文本，保留 pPr/rPr/drawing
 *   5. 用 store 方式重新打包全部部件（保持 zip 合法）
 *
 * 若对齐大面积失败（换结构、扫描件等），返回 {ok:false}，调用方回退到旧的
 * 「重新排版」生成器，保证永远能导出可用文件。
 */
'use strict';
const { unzip } = require('./source-style');

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function unesc(s) {
  return String(s == null ? '' : s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

// ---------- 极简 ZIP (store) 打包 ----------
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
function zipStore(files) {
  var chunks = [], central = [], offset = 0;
  files.forEach(function (f) {
    var nameBuf = Buffer.from(f.name, 'utf8');
    var data = f.data, crc = crc32(data);
    var local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10); local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18); local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26); local.writeUInt16LE(0, 28);
    chunks.push(local, nameBuf, data);
    var cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0x0800, 8);
    cd.writeUInt16LE(0, 10);
    cd.writeUInt16LE(0, 12); cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(data.length, 20); cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt16LE(0, 30); cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34); cd.writeUInt16LE(0, 36);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, nameBuf);
    offset += local.length + nameBuf.length + data.length;
  });
  var centralSize = central.reduce(function (s, b) { return s + b.length; }, 0);
  var centralOffset = offset;
  var eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4); eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralSize, 12); eocd.writeUInt32LE(centralOffset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat(chunks.concat(central, [eocd]));
}

/**
 * 接受所有修订（accept all changes）：
 *  - 删除 <w:del>...</w:del> 整块（含其中的 <w:delText>，即灰色删除文字）
 *  - 解开 <w:ins>...</w:ins> 外壳，保留其中的 run/<w:t>（即插入的新文字）
 *  - 移除格式变更标记 <w:rPrChange>/<w:pPrChange>/<w:tblPrChange>/<w:sectPrChange> 等
 *  - 移除批注锚点 <w:commentRangeStart>/<w:commentRangeEnd>/<w:commentReference>
 * 这样做的原因：原简历若带修订痕迹，Word 会同时显示"删除的旧文字(灰色)+插入的新文字"，
 * 保真替换后旧文字残留 → 截图里的文字重叠 + 左侧修订/批注锚点图标。
 * 用基于栈的扫描器处理，可正确应对 <w:ins> 内嵌 <w:del> 等嵌套情况。
 */
function acceptAllRevisions(xml) {
  // 先处理 del / ins 成对块（大小写固定为 w: 前缀，docx 规范）
  var OUTER = /<w:(del|ins|moveFrom|moveTo)\b([^>]*)>/g;
  var m, out = '', last = 0, openTag;
  function findClose(src, from, name) {
    // 返回 <w:name ...> 与之配对的 </w:name> 位置（考虑同名嵌套）
    var depth = 1, re = new RegExp('<(/?)w:' + name + '\\b[^>]*>', 'g');
    re.lastIndex = from; // 开标签已跳过，depth 从 1 开始计
    var t;
    while ((t = re.exec(src)) !== null) {
      if (t[1] === '/') { if (depth === 1) return { start: t.index, end: re.lastIndex }; depth--; }
      else depth++;
    }
    return null;
  }
  var guard = 0;
  OUTER.lastIndex = 0;
  while ((m = OUTER.exec(xml)) !== null && guard++ < 5000) {
    var name = m[1], openStart = m.index, openEnd = OUTER.lastIndex;
    var close = findClose(xml, openEnd, name);
    if (!close) continue; // 配对失败则保留原样，避免破坏结构
    out += xml.slice(last, openStart);
    var inner = xml.slice(openEnd, close.start);
    if (name === 'del' || name === 'moveFrom') {
      // 删除块整体丢弃（内含 delText）
    } else {
      // 插入块：保留内部内容（去掉可能的 delText，保留 w:t）
      out += inner.replace(/<w:delText\b([^>]*)>/g, '<w:t$1>').replace(/<\/w:delText>/g, '</w:t>');
    }
    last = close.end;
    OUTER.lastIndex = close.end;
  }
  out += xml.slice(last);
  // 多遍剥离安全网：处理「<w:ins> 内嵌 <w:del>」等嵌套——主扫描器跳过了 ins 内部，
  // 可能残留 <w:del><w:t>...</w:t></w:del>。非贪婪逐层剥离，最多 4 遍直到稳定。
  for (var pass = 0; pass < 4; pass++) {
    var before = out;
    out = out
      .replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, '')       // 删除块（含内部文字）整块丢弃
      .replace(/<w:moveFrom\b[^>]*>[\s\S]*?<\/w:moveFrom>/g, '')
      .replace(/<w:del\b[^>]*\/>/g, '')
      .replace(/<w:moveFrom\b[^>]*\/>/g, '')
      .replace(/<w:ins\b([^>]*)>([\s\S]*?)<\/w:ins>/g, '$2') // 插入块：解开外壳保留内容
      .replace(/<w:moveTo\b([^>]*)>([\s\S]*?)<\/w:moveTo>/g, '$2');
    if (out === before) break;
  }
  // 移除孤立的 change / comment 标记（多为自闭合或带内容的小元素）
  out = out
    .replace(/<w:(?:rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange|numberingChange)\b[^>]*>[\s\S]*?<\/w:\1>/g, '')
    .replace(/<w:(?:rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|sectPrChange|numberingChange)\b[^>]*\/>/g, '')
    .replace(/<w:commentRange(?:Start|End)\b[^>]*\/>/g, '')
    .replace(/<w:commentReference\b[^>]*\/>/g, '')
    .replace(/<w:annotationRef\b[^>]*\/>/g, '');
  // delText 即使在孤立残留里也转为普通文本，避免 Word 当删除文字渲染
  out = out.replace(/<w:delText\b([^>]*)>/g, '<w:t$1>').replace(/<\/w:delText>/g, '</w:t>');
  return out;
}

// ---------- 段落解析 ----------
// 注意：w:p 不会嵌套 w:p；<w:pPr> 因 \b 不会误匹配
var PARA_RE = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
var T_RE = /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g;

function paraText(inner) {
  var out = '', m;
  T_RE.lastIndex = 0;
  while ((m = T_RE.exec(inner)) !== null) out += unesc(m[2]);
  return out;
}

/** 去掉行首的 markdown / 项目符号，避免替换进 docx 后出现双重符号 */
function cleanLine(line) {
  return String(line || '')
    .replace(/^\s*#{1,6}\s*/, '')
    .replace(/^\s*[-*•◦·☐☑✅▪►▶]\s*/, '')
    .replace(/^\s*\d+[.)、]\s*/, '')
    .replace(/\*\*/g, '').replace(/__/g, '').replace(/`/g, '')
    .trim();
}

/** 字符二元组相似度（对称 Dice 系数，恒在 [0,1]），用于把润色行对齐回原段落 */
function similarity(a, b) {
  a = String(a || '').replace(/\s/g, '');
  b = String(b || '').replace(/\s/g, '');
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.length < 2 || b.length < 2) return a.indexOf(b) >= 0 || b.indexOf(a) >= 0 ? 0.5 : 0;
  var ga = {}, gb = {};
  var i, g;
  for (i = 0; i < a.length - 1; i++) { g = a.substr(i, 2); ga[g] = (ga[g] || 0) + 1; }
  for (i = 0; i < b.length - 1; i++) { g = b.substr(i, 2); gb[g] = (gb[g] || 0) + 1; }
  // 交集按每元组较小计数求和，分母为两边二元组总数 —— Dice 系数 ≤ 1
  var inter = 0, na = 0, nb = 0;
  for (g in ga) { na += ga[g]; if (gb[g]) inter += Math.min(ga[g], gb[g]); }
  for (g in gb) nb += gb[g];
  return (na + nb) ? (2 * inter) / (na + nb) : 0;
}

/** 把某段的所有 <w:t>：第一段填新文字，其余清空（保留 run / rPr 结构） */
function replaceParaText(inner, newText) {
  var first = true;
  return inner.replace(T_RE, function (m, open, content, close) {
    if (first) { first = false; return open + esc(newText) + close; }
    return open + close;
  });
}

/**
 * 主函数：保真替换导出
 * @param {Buffer} originalDocx 原始 docx 字节
 * @param {string} polishedText 润色后的简历正文（可能含 markdown）
 * @returns {Buffer|null} 成功返回 docx Buffer；无法保真时返回 null
 */
function preserveReplace(originalDocx, polishedText) {
  try {
    const files = unzip(originalDocx);
    if (!files['word/document.xml']) return null;
    // 先接受所有修订：清除灰色删除文字与插入/批注标记，避免导出后文字重叠
    var rawDocXml = files['word/document.xml'].toString('utf8');
    const docXml = acceptAllRevisions(rawDocXml);
    files['word/document.xml'] = Buffer.from(docXml, 'utf8');
    // 关闭修订模式，否则 Word 打开后可能再次把新写入的文字标成修订
    if (files['word/settings.xml']) {
      var settings = files['word/settings.xml'].toString('utf8')
        .replace(/<w:trackChanges\b[^>]*\/>/g, '')
        .replace(/<w:doNotTrackMoves\b[^>]*\/>/g, '');
      files['word/settings.xml'] = Buffer.from(settings, 'utf8');
    }

    // 原文段落（只处理含文字的段落；含图片的段落整体保留）
    var paras = [], pm;
    PARA_RE.lastIndex = 0;
    while ((pm = PARA_RE.exec(docXml)) !== null) {
      var inner = pm[1];
      var text = paraText(inner);
      var hasT = /<w:t\b/.test(inner);
      var hasImg = /<w:drawing|<w:pict/.test(inner);
      paras.push({ full: pm[0], inner: inner, text: text, hasT: hasT, hasImg: hasImg });
    }
    var textParas = paras.filter(function (p) { return p.hasT && p.text.trim(); });
    if (textParas.length < 2) return null; // 结构太少，放弃保真

    // 润色正文按行拆分、清洗
    var lines = String(polishedText || '').replace(/\r/g, '').split('\n')
      .map(cleanLine).filter(function (l) { return l.length > 0; });
    if (lines.length < 2) return null;

    // ===== 按板块分块匹配 =====
    // 原件是表格排版的复杂简历，润色常重排板块（项目经验提前、新增专业技能）。
    // 若做全局匹配会把「自我评价」塞进「校外经历」，产生错位/重复/半截句。
    // 正确做法：先给每个原段落、每条润色行标注所属板块，只允许「同一板块内」一对一匹配。
    var SECTIONS = ['个人信息', '教育背景', '校外经历', '项目经验', '实习经历', '工作经历', '荣誉', '奖励', '专业技能', '技能', '自我评价'];
    function sectionOf(s) {
      s = String(s || '').replace(/\s/g, '');
      if (!s || s.length > 10) return null;
      for (var k = 0; k < SECTIONS.length; k++) if (s.indexOf(SECTIONS[k]) >= 0) return SECTIONS[k];
      return null;
    }
    // 顺序标注：遇到板块标题就切换 current，其余行继承当前板块；头部为 __head__
    function annotate(items, getText) {
      var cur = '__head__';
      return items.map(function (it) {
        var s = sectionOf(getText(it));
        if (s) cur = s;
        return cur;
      });
    }
    var paraSec = annotate(textParas, function (p) { return p.text; });
    var lineSec = annotate(lines, function (l) { return cleanLine(l); });

    // 原件存在的板块集合（用于区分「润色新增板块」与「已存在板块的重复行」）
    var srcSectionSet = {};
    paraSec.forEach(function (s) { if (s !== '__head__') srcSectionSet[s] = true; });

    // 板块内一对一匹配（Dice 相似度，阈值放宽到 0.34 以覆盖措辞大幅改写）。
    // 仅「含图片的段落」排除匹配，确保证件照 100% 原样；其余段落（含抬头的姓名/求职意向
    // 等 __head__ 文字段）照常参与——它们虽无板块标题，但同属 __head__，彼此匹配即可。
    var pairs = [];
    for (var pi = 0; pi < textParas.length; pi++) {
      if (textParas[pi].hasImg) continue;
      for (var li = 0; li < lines.length; li++) {
        if (lineSec[li] !== paraSec[pi]) continue;
        var sc = similarity(textParas[pi].text, lines[li]);
        if (sc >= 0.34) pairs.push({ p: pi, l: li, s: sc });
      }
    }
    pairs.sort(function (a, b) { return b.s - a.s; });
    var usedLine = new Array(lines.length).fill(false);
    var usedPara = new Array(textParas.length).fill(false);
    var matchCount = 0;
    pairs.forEach(function (pr) {
      if (usedLine[pr.l] || usedPara[pr.p]) return;
      var nt = lines[pr.l];
      // 原件常把「标签：」和「值」拆成两个相邻文字段（如「籍贯：」「四川宜宾」）。
      // 润色行是「标签：值」合并写法。若本段是值段（前一段为纯标签），只取冒号后的值，
      // 避免「籍贯：籍贯：四川宜宾」这类重复。
      var pp = pr.p;
      if (pp > 0) {
        var prev = String(textParas[pp - 1].text || '').replace(/\s/g, '');
        if (/[:：]$/.test(prev) && prev.length <= 8 && paraSec[pp - 1] === paraSec[pp]) {
          var cm = nt.match(/^[^:：]{1,8}[:：]\s*(.+)$/);
          if (cm) nt = cm[1];
        }
      }
      textParas[pr.p].newText = nt;
      usedLine[pr.l] = true; usedPara[pr.p] = true; matchCount++;
    });

    // 对齐质量门槛：分母只算参与匹配的段落/行（排除含证件照的图片段，它不参与替换）。
    var candidateParas = textParas.filter(function (p) { return !p.hasImg; }).length;
    var matchRate = matchCount / Math.max(candidateParas, lines.length, 1);
    if (matchRate < 0.45) return null;

    // 重建 document.xml：逐段替换
    var rebuilt = docXml;
    paras.forEach(function (p) {
      var newInner = (p.hasT && p.newText) ? replaceParaText(p.inner, p.newText) : p.inner;
      var newFull = p.full.replace(p.inner, newInner);
      rebuilt = rebuilt.replace(p.full, newFull);
    });

    // 只把「润色真正新增的板块」（原件没有，如专业技能/自我评价）追加到表格后，
    // 作为普通段落。已存在板块里没匹配上的行一律丢弃（避免文末堆积重复内容）。
    var leftovers = [];
    lines.forEach(function (t, i) {
      if (usedLine[i]) return;
      var sec = lineSec[i];
      var isSectionTitle = !!sectionOf(cleanLine(t));
      if (sec !== '__head__' && !srcSectionSet[sec]) leftovers.push(t);
      else if (isSectionTitle && !srcSectionSet[sectionOf(cleanLine(t))]) leftovers.push(t);
    });
    if (leftovers.length) {
      // 板块标题用加粗，其余普通段落，视觉上自成一体、不突兀
      var addition = leftovers.map(function (t) {
        var bold = sectionOf(cleanLine(t)) ? '<w:b/>' : '';
        return '<w:p><w:pPr><w:spacing w:before="80" w:after="60"/></w:pPr>' +
          '<w:r><w:rPr>' + bold + '</w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r></w:p>';
      }).join('');
      rebuilt = rebuilt.replace(/<w:sectPr\b/, addition + '<w:sectPr');
    }

    // 健全性检查：确保关键个人信息仍在、且没有把两段替换成同一句（重复）。
    // 任一失败则回退重新生成器，避免产出错位/重复的简历。
    var finalTexts = [];
    PARA_RE.lastIndex = 0;
    var fm;
    while ((fm = PARA_RE.exec(rebuilt)) !== null) finalTexts.push(paraText(fm[1]).replace(/\s/g, ''));
    var joined = finalTexts.join('\n');
    // 关键个人信息锚点（从润色正文里提取手机号/邮箱/姓名/学校等）必须都出现在结果中
    var anchors = [];
    var phoneM = text.match(/1\d{10}/); if (phoneM) anchors.push(phoneM[0]);
    var mailM = text.match(/[\w.\-]+@[\w.\-]+\.\w+/); if (mailM) anchors.push(mailM[0]);
    for (var ai = 0; ai < anchors.length; ai++) {
      if (joined.indexOf(anchors[ai]) < 0) return null; // 个人信息丢失 → 回退
    }
    // 检测重复行：同一段润色文字被塞进多个原段落（理想情况全局唯一分配已避免，这里兜底）
    var seen = {};
    for (var fi = 0; fi < finalTexts.length; fi++) {
      var ft = finalTexts[fi];
      if (ft.length >= 8) {
        if (seen[ft]) return null; // 出现整段重复 → 回退
        seen[ft] = true;
      }
    }

    files['word/document.xml'] = Buffer.from(rebuilt, 'utf8');

    // 重新打包全部部件（顺序尽量保持原样）
    var outFiles = Object.keys(files).map(function (name) {
      return { name: name, data: files[name] };
    });
    return zipStore(outFiles);
  } catch (e) {
    return null;
  }
}

/** 从原 docx 提取所有图片（供 PDF 嵌入证件照），返回 [{name,data,ext}] */
function extractImages(originalDocx) {
  try {
    const files = unzip(originalDocx);
    var imgs = [];
    Object.keys(files).forEach(function (name) {
      if (/^word\/media\//i.test(name) && /\.(png|jpe?g|gif|bmp|webp)$/i.test(name)) {
        var m = name.match(/\.(png|jpe?g|gif|bmp|webp)$/i);
        imgs.push({ name: name, data: files[name], ext: (m ? m[1].toLowerCase() : 'png') });
      }
    });
    return imgs;
  } catch (e) { return []; }
}

module.exports = { preserveReplace, extractImages, zipStore, crc32, esc };
