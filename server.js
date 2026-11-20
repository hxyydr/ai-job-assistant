/**
 * AI 求职小助手 - 本地服务
 * 启动：node server.js   然后浏览器打开 http://localhost:3000
 * 零第三方依赖，仅使用 Node 18+ 内置的 http / fetch。
 *
 * 大模型接入：任何 OpenAI Chat Completions 兼容接口均可，
 * 例如 DeepSeek / 通义千问 compatible-mode / Moonshot / OpenAI。
 * 在设置面板填写接口地址、密钥、模型名即可启用；不填则使用内置本地规则引擎。
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { generateDocx } = require('./lib/docx');
const { generatePdf } = require('./lib/pdf');
const { extractStyle, resolveFontPath, extractPdfImages } = require('./lib/source-style');
const { preserveReplace, extractImages } = require('./lib/docx-preserve');

// 去掉润色正文末尾混入的「对话式收尾话术」——这些是聊天话术，绝不能进导出的简历文件。
// 典型：'待你补充的信息（原文确实缺失）：''补完后可直接使用。需要我…面试题库/模拟面试吗？'
//      '提醒：【】中内容…''说明：诊断仅基于…''如需…欢迎告诉我' 等。
function stripChatter(s) {
  var lines = String(s || '').split('\n');
  var CHATTER = [
    /待你补充的信息/,
    /原文确实缺失/,
    /补完后?可?直接?使用/,
    /补完后可直接使用/,
    /需要我(针对|帮你|给你)?[^\n]{0,30}(面试|题库|模拟面试)/,
    /是否需要我[^\n]{0,30}(面试|题库)/,
    /提醒[：:]\s*【】/,
    /【】中内容[^\n]{0,30}(填写|真实|切勿)/,
    /说明[：:]\s*诊断仅基于/,
    /润色不能替代你的思考/,
    /以上(内容|建议)[^\n]{0,20}(参考|采纳)/,
    /以个人思考为主/,
    /不要直接复制AI内容投递/,
    /^---\s*$/ // 收尾分隔线；正文中的分隔线通常在此之后，一律截掉其后内容
  ];
  var cut = -1;
  for (var i = 0; i < lines.length; i++) {
    var t = lines[i].trim();
    if (!t) continue;
    for (var k = 0; k < CHATTER.length; k++) {
      if (CHATTER[k].test(t)) { cut = i; break; }
    }
    if (cut >= 0) break;
  }
  if (cut >= 0) lines = lines.slice(0, cut);
  return lines.join('\n');
}

// 清洗导出正文：前端已拆过标记，这里再防御性去掉残留的【改动说明】/【润色后简历正文】标记与说明段
function cleanExportText(s) {
  s = String(s || '').slice(0, 100000).replace(/\r/g, '');
  const m = s.match(/【润色后简历正文】/);
  if (m) s = s.slice(m.index + m[0].length);
  // 若仍残留【改动说明】整段（模型未分隔时的兜底），截掉标记本身这行
  s = s.replace(/^\s*【润色后简历正文】\s*[\r\n]?/gm, '');
  s = stripChatter(s);
  return s.replace(/^\s*[\r\n]+/, '').trim().replace(/\s+$/, '');
}

// 把前端可能传来的扩展名别名归一：word/doc -> docx
function normalizeExt(ext) {
  const e = String(ext || '').toLowerCase().replace(/^\./, '');
  if (e === 'word' || e === 'doc') return 'docx';
  if (e === 'pdf') return 'pdf';
  return e;
}

// 从导出请求里解析原始简历样式（若有）；返回 {style, ok, ext}
function styleFromBody(body) {
  if (!body || !body.sourceBase64 || !body.sourceExt) return { style: null, ok: false };
  try {
    const buf = Buffer.from(String(body.sourceBase64), 'base64');
    if (buf.length < 100 || buf.length > 30 * 1024 * 1024) return { style: null, ok: false };
    const style = extractStyle(buf, normalizeExt(body.sourceExt));
    if (style && style.ok) {
      // PDF 导出：把字体族解析成实际字体文件路径（docx 字体名 → 系统 ttf/ttc）
      const fam = style.family || style.font;
      if (fam) style.fontPath = resolveFontPath(fam);
    }
    return { style: style, ok: !!(style && style.ok) };
  } catch (e) { return { style: null, ok: false }; }
}

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, 'public');
const CONFIG_PATH = path.join(ROOT, 'config.json');

// 数据持久化目录：使用记录与用户建议
const DATA_DIR = path.join(ROOT, '对话记录与建议');
const RECORDS_DIR = path.join(DATA_DIR, '对话记录');      // 对话记录，按天存为 .jsonl
const FEEDBACK_DIR = path.join(DATA_DIR, '建议');         // 用户建议，存为 feedback.json
const FEEDBACK_FILE = path.join(FEEDBACK_DIR, 'feedback.json');

// 启动时确保目录与文件存在
function ensureDataDirs() {
  try {
    [DATA_DIR, RECORDS_DIR, FEEDBACK_DIR].forEach(function (d) {
      if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
    });
    if (!fs.existsSync(FEEDBACK_FILE)) fs.writeFileSync(FEEDBACK_FILE, '[]', 'utf8');
  } catch (e) {
    console.error('初始化数据目录失败：', e.message);
  }
}
ensureDataDirs();

function todayKey() {
  const d = new Date();
  const p = function (n) { return String(n).padStart(2, '0'); };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

function appendRecord(rec) {
  const file = path.join(RECORDS_DIR, todayKey() + '.jsonl');
  const line = JSON.stringify(Object.assign({ time: new Date().toISOString() }, rec)) + '\n';
  fs.appendFileSync(file, line, 'utf8');
}

function loadFeedback() {
  try {
    const arr = JSON.parse(fs.readFileSync(FEEDBACK_FILE, 'utf8'));
    return Array.isArray(arr) ? arr : [];
  } catch (e) { return []; }
}

function saveFeedback(arr) {
  fs.writeFileSync(FEEDBACK_FILE, JSON.stringify(arr, null, 2), 'utf8');
}

// 读取全部记录行（最新在前），供统计与管理查看共用
function collectRecords() {
  var lines = [];
  try {
    fs.readdirSync(RECORDS_DIR).filter(function (f) { return /\.jsonl$/.test(f); }).sort().reverse()
      .forEach(function (f) {
        var content = fs.readFileSync(path.join(RECORDS_DIR, f), 'utf8').trim();
        if (content) lines = lines.concat(content.split('\n'));
      });
  } catch (e) {}
  return lines.filter(Boolean).map(function (l) {
    try { return JSON.parse(l); } catch (e) { return null; }
  }).filter(Boolean);
}

// 公开统计：只包含聚合数字，绝不包含任何对话内容或用户身份
function computeStats() {
  var all = collectRecords();
  var users = {};
  var days = {};
  var conversations = 0;
  all.forEach(function (r) {
    if (r.user) users[r.user] = 1;
    if (r.role === 'user') conversations++;
    if (r.time) days[String(r.time).slice(0, 10)] = 1;
  });
  return {
    ok: true,
    messages: all.length,
    conversations: conversations,
    visitors: Object.keys(users).length,
    days: Object.keys(days).length,
    feedback: loadFeedback().length
  };
}

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.pdf': 'application/pdf'
};

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => {
      data += chunk;
      if (data.length > 5 * 1024 * 1024) {
        req.destroy();
        reject(new Error('请求体过大'));
      }
    });
    req.on('end', () => {
      try { resolve(JSON.parse(data || '{}')); }
      catch (e) { reject(new Error('请求体不是合法 JSON')); }
    });
    req.on('error', reject);
  });
}

function sendJSON(res, code, obj) {
  res.writeHead(code, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS'
  });
  res.end(JSON.stringify(obj));
}

function readConfig() {
  var cfg = {};
  try { cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); }
  catch (e) {}
  // 环境变量优先级更高，便于不落盘地注入密钥
  if (process.env.LLM_BASE_URL) cfg.baseUrl = process.env.LLM_BASE_URL;
  if (process.env.LLM_API_KEY) cfg.apiKey = process.env.LLM_API_KEY;
  if (process.env.LLM_MODEL) cfg.model = process.env.LLM_MODEL;
  return cfg;
}

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') return sendJSON(res, 204, {});

  let url;
  try { url = new URL(req.url, `http://${req.headers.host || 'localhost'}`); }
  catch (e) { return sendJSON(res, 400, { error: 'BAD_URL' }); }

  // 健康检查
  if (url.pathname === '/api/health') {
    return sendJSON(res, 200, { ok: true, time: new Date().toISOString() });
  }

  // 服务端配置状态（不返回密钥）
  if (url.pathname === '/api/config') {
    const c = readConfig();
    return sendJSON(res, 200, {
      configured: !!(c.apiKey && c.baseUrl && c.model),
      baseUrl: c.baseUrl || '',
      model: c.model || ''
    });
  }

  // 上报一条对话记录（写入 对话记录与建议/对话记录/YYYY-MM-DD.jsonl）
  if (url.pathname === '/api/records' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      const rec = {
        user: String(body.user || '匿名用户').slice(0, 60),
        mode: String(body.mode || '').slice(0, 40),
        role: String(body.role || '').slice(0, 20),
        content: String(body.content || '').slice(0, 20000)
      };
      appendRecord(rec);
      return sendJSON(res, 200, { ok: true });
    } catch (e) {
      return sendJSON(res, 500, { ok: false, error: String((e && e.message) || e) });
    }
  }

  // 公开统计：任何人可见，只返回数字，不含任何对话内容或用户身份
  if (url.pathname === '/api/stats' && req.method === 'GET') {
    return sendJSON(res, 200, computeStats());
  }

  // 查看完整记录（仅管理员）：需提供正确口令，否则拒绝
  if (url.pathname === '/api/records' && req.method === 'GET') {
    const c = readConfig();
    const token = String(url.searchParams.get('token') || req.headers['x-admin-token'] || '');
    if (!c.adminToken || token !== String(c.adminToken)) {
      return sendJSON(res, 401, { ok: false, error: 'UNAUTHORIZED', message: '口令不正确，无法查看记录。' });
    }
    try {
      const limit = Math.min(Number(url.searchParams.get('limit') || 100), 1000);
      var all = collectRecords();
      return sendJSON(res, 200, { ok: true, count: all.length, items: all.slice(-limit).reverse() });
    } catch (e) {
      return sendJSON(res, 500, { ok: false, error: String((e && e.message) || e) });
    }
  }

  // 提交用户建议（写入 对话记录与建议/建议/feedback.json）
  if (url.pathname === '/api/feedback' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      const content = String(body.content || '').trim();
      if (!content) return sendJSON(res, 400, { ok: false, error: '建议内容不能为空' });
      var arr = loadFeedback();
      arr.push({
        id: Date.now() + '-' + Math.random().toString(36).slice(2, 8),
        user: String(body.user || '匿名用户').slice(0, 60),
        content: content.slice(0, 5000),
        time: new Date().toISOString()
      });
      saveFeedback(arr);
      return sendJSON(res, 200, { ok: true, count: arr.length });
    } catch (e) {
      return sendJSON(res, 500, { ok: false, error: String((e && e.message) || e) });
    }
  }

  // 获取建议列表（建议墙展示）
  if (url.pathname === '/api/feedback' && req.method === 'GET') {
    return sendJSON(res, 200, { ok: true, items: loadFeedback().slice().reverse() });
  }

  // 导出润色结果为 .docx（零依赖生成标准 Word 文档）
  if (url.pathname === '/api/export-docx' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      var text = cleanExportText(String(body.text || ''));
      if (!text.trim()) return sendJSON(res, 400, { ok: false, error: '内容为空' });
      var src = styleFromBody(body);
      var title = src.ok ? (body.sourceName ? String(body.sourceName).replace(/\.(docx?|pdf)$/i, '') + ' · 润色版' : '简历 · 润色版')
                         : 'AI 润色简历 · ' + (body.mode === 'polish' ? '优化版' : '文档');
      // 保真优先：若上传的是原始 docx，直接在其上替换文字，完整保留证件照/复选框/样式/排版
      var buf = null;
      var origBuf = null;
      if (src.ok && normalizeExt(body.sourceExt) === 'docx') {
        try {
          origBuf = Buffer.from(String(body.sourceBase64), 'base64');
          buf = preserveReplace(origBuf, text);
        } catch (e) { buf = null; }
      }
      if (!buf) {
        // 回退重新生成：保留原文件字体/字号/颜色/页边距，并把原证件照重新嵌入
        var images = [];
        if (origBuf) {
          try { images = extractImages(origBuf); } catch (e) { images = []; }
        }
        buf = generateDocx(text, title, src.ok ? src.style : null, { customTitle: src.ok, images: images });
      }
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="resume-polished.docx"',
        'Content-Length': buf.length,
        'Access-Control-Allow-Origin': '*'
      });
      return res.end(buf);
    } catch (e) {
      return sendJSON(res, 500, { ok: false, error: String((e && e.message) || e) });
    }
  }

  // 导出为 .pdf（零依赖生成中文 PDF，文字可选可搜索）
  if (url.pathname === '/api/export-pdf' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      var text = cleanExportText(String(body.text || ''));
      if (!text.trim()) return sendJSON(res, 400, { ok: false, error: '内容为空' });
      var src = styleFromBody(body);
      var title = src.ok ? (body.sourceName ? String(body.sourceName).replace(/\.(docx?|pdf)$/i, '') + ' · 润色版' : '简历 · 润色版') : 'AI 润色简历 · ' + (
        body.mode === 'polish' ? '优化版' :
        body.mode === 'ppt' ? 'PPT 大纲' :
        body.mode === 'summary' ? '文档总结' : '文档'
      );
      // 从原文件提取证件照嵌入 PDF（docx 用 media，pdf 用内嵌图片）
      var images = [];
      if (src.ok && body.sourceBase64) {
        try {
          var srcBuf = Buffer.from(String(body.sourceBase64), 'base64');
          var normExt = normalizeExt(body.sourceExt);
          if (normExt === 'docx') images = extractImages(srcBuf);
          else if (normExt === 'pdf') images = extractPdfImages(srcBuf);
        } catch (e) { images = []; }
      }
      var buf = generatePdf(text, title, src.ok ? src.style : null, { customTitle: src.ok, images: images });
      res.writeHead(200, {
        'Content-Type': 'application/pdf',
        'Content-Disposition': 'attachment; filename="resume-export.pdf"',
        'Content-Length': buf.length,
        'Access-Control-Allow-Origin': '*'
      });
      return res.end(buf);
    } catch (e) {
      return sendJSON(res, 500, { ok: false, error: String((e && e.message) || e) });
    }
  }

  // 大模型代理
  if (url.pathname === '/api/chat' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      const cfg = Object.assign({}, readConfig(), body.config || {});
      if (!cfg.apiKey || !cfg.baseUrl || !cfg.model) {
        return sendJSON(res, 400, {
          error: 'NO_CONFIG',
          message: '未配置大模型：请填写接口地址、密钥与模型名，或使用本地规则引擎。'
        });
      }

      const base = String(cfg.baseUrl).replace(/\/+$/, '');
      const endpoint = /\/chat\/completions$/.test(base) ? base : base + '/chat/completions';
      const payload = {
        model: cfg.model,
        messages: [
          { role: 'system', content: String(body.system || '') },
          ...(Array.isArray(body.messages) ? body.messages : [])
        ],
        temperature: typeof body.temperature === 'number' ? body.temperature : 0.7,
        stream: false
      };

      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 90000);
      let r;
      try {
        r = await fetch(endpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + cfg.apiKey
          },
          body: JSON.stringify(payload),
          signal: controller.signal
        });
      } finally {
        clearTimeout(timer);
      }

      const text = await r.text();
      let data;
      try { data = JSON.parse(text); }
      catch (e) { throw new Error('模型返回非 JSON 内容：' + text.slice(0, 200)); }

      if (!r.ok) {
        return sendJSON(res, r.status, {
          error: 'API_ERROR',
          message: (data && data.error && data.error.message) || text.slice(0, 300)
        });
      }
      const content = data && data.choices && data.choices[0] && data.choices[0].message
        && data.choices[0].message.content;
      if (!content) return sendJSON(res, 502, { error: 'EMPTY', message: '模型返回内容为空' });
      return sendJSON(res, 200, { content });
    } catch (e) {
      return sendJSON(res, 500, { error: 'EXCEPTION', message: String((e && e.message) || e) });
    }
  }

  // 静态文件
  let pathname = decodeURIComponent(url.pathname);
  if (pathname === '/' || pathname === '') pathname = '/index.html';
  let filePath = path.join(PUBLIC_DIR, pathname.replace(/^\/+/, ''));
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    return res.end('Forbidden');
  }
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    filePath = path.join(PUBLIC_DIR, 'index.html');
  }
  res.writeHead(200, {
    'Content-Type': MIME[path.extname(filePath)] || 'application/octet-stream',
    'Cache-Control': 'no-cache, no-store, must-revalidate',
    'Pragma': 'no-cache',
    'Expires': '0'
  });
  fs.createReadStream(filePath).pipe(res);
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('AI 求职小助手已启动：http://0.0.0.0:' + PORT);
});
