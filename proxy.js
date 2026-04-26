// 轻量 Polymarket 代理 + 静态文件服务器（预览用）
// 用法：node proxy.js
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const PORT = 3002;
const POLYMARKET_BASE = 'https://gamma-api.polymarket.com';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript',
  '.css':  'text/css',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg':  'image/svg+xml',
  '.json': 'application/json',
  '.csv':  'text/csv',
  '.zip':  'application/zip',
};

const LOG_DIR = path.join(__dirname, 'logs');
const USAGE_LOG = path.join(LOG_DIR, 'usage.jsonl');
const DAILY_QUOTA = Number(process.env.DAILY_QUOTA || 30);
const OUTPUT_ROOT = path.join(__dirname, 'output');

if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR, { recursive: true });

let isGenerating = false;

function todayKey() {
  return new Date().toISOString().slice(0, 10);
}

function getClientIP(req) {
  const xff = req.headers['x-forwarded-for'];
  if (xff) return String(xff).split(',')[0].trim();
  return (req.socket.remoteAddress || '').replace(/^::ffff:/, '');
}

function appendUsage(entry) {
  try {
    fs.appendFileSync(USAGE_LOG, JSON.stringify(entry) + '\n');
  } catch (err) {
    console.error('写入 usage.jsonl 失败：', err.message);
  }
}

function readUsageLog() {
  if (!fs.existsSync(USAGE_LOG)) return [];
  return fs.readFileSync(USAGE_LOG, 'utf-8')
    .split('\n')
    .filter(Boolean)
    .map(line => { try { return JSON.parse(line); } catch { return null; } })
    .filter(Boolean);
}

function getTodayStartCount() {
  const today = todayKey();
  return readUsageLog().filter(e => (e.ts || '').startsWith(today) && e.event === 'start').length;
}

function proxyPolymarket(req, res, apiPath) {
  const url = `${POLYMARKET_BASE}${apiPath}`;
  https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, (apiRes) => {
    res.writeHead(apiRes.statusCode, {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*',
    });
    apiRes.pipe(res);
  }).on('error', (err) => {
    res.writeHead(502);
    res.end(JSON.stringify({ error: err.message }));
  });
}

function handleGenerateStream(req, res) {
  const urlObj = new URL(req.url, `http://localhost:${PORT}`);
  const templateKey = urlObj.searchParams.get('template');
  const ip = getClientIP(req);

  if (!templateKey) {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ error: '缺少 template 参数' }));
  }

  if (isGenerating) {
    res.writeHead(409, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ error: '正在生成中，请稍候' }));
  }

  const todayCount = getTodayStartCount();
  if (todayCount >= DAILY_QUOTA) {
    res.writeHead(429, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ error: `今日额度已用完（${DAILY_QUOTA} 次/天），明天再试` }));
  }

  isGenerating = true;
  appendUsage({ ts: new Date().toISOString(), event: 'start', template: templateKey, ip });

  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*',
  });

  const sendEvent = (type, data) => {
    if (type === 'message') {
      res.write(`data: ${data}\n\n`);
    } else {
      res.write(`event: ${type}\ndata: ${data}\n\n`);
    }
  };

  sendEvent('message', `开始生成模板：${templateKey}（今日第 ${todayCount + 1}/${DAILY_QUOTA} 次）`);

  let zipFilename = null;
  let outputDir = null;

  const child = spawn('node', ['generate.js', '--template', templateKey], {
    cwd: __dirname,
    env: { ...process.env, FORCE_COLOR: '0' },
  });

  const onLine = (chunk) => {
    const lines = chunk.toString().split('\n');
    lines.forEach(line => {
      const trimmed = line.replace(/\r/g, '');
      if (!trimmed.trim()) return;
      const zipMatch = trimmed.match(/📦\s+(\S+\.zip)/);
      if (zipMatch) zipFilename = zipMatch[1];
      const dirMatch = trimmed.match(/所有图片已保存到[：:]\s*(.+?)\s*$/);
      if (dirMatch) outputDir = dirMatch[1];
      sendEvent('message', trimmed);
    });
  };

  child.stdout.on('data', onLine);
  child.stderr.on('data', onLine);

  child.on('close', (code) => {
    isGenerating = false;
    let zipRel = null;
    if (code === 0 && outputDir && zipFilename) {
      const abs = path.join(outputDir, zipFilename);
      if (fs.existsSync(abs)) zipRel = path.relative(__dirname, abs);
    }
    appendUsage({
      ts: new Date().toISOString(),
      event: 'end',
      template: templateKey,
      ip,
      status: code === 0 ? 'success' : 'fail',
      exitCode: code,
      zipPath: zipRel,
    });
    if (code === 0) {
      sendEvent('done', JSON.stringify({ zipPath: zipRel, filename: zipFilename }));
    } else {
      sendEvent('error', `生成失败（退出码 ${code}）`);
    }
    res.end();
  });

  req.on('close', () => {
    if (isGenerating) {
      child.kill();
      isGenerating = false;
    }
  });
}

const PREVIEW_BG_DIRS = {
  classic:       'backgrounds-NBA',
  comprehensive: 'backgrounds-NBA',
  worldcup:      'backgrounds- football',
  football:      'backgrounds- football',
  coinprice:     'backgrounds-global',
  global:        'backgrounds-global',
  f1:            'backgrounds-F1',
  f1driver:      'backgrounds-F1',
};

function pickBgFile(bgDir) {
  if (!fs.existsSync(bgDir)) return null;
  const files = fs.readdirSync(bgDir).filter(f => /\.(png|jpg)$/i.test(f));
  if (!files.length) return null;
  const preferred = ['zh-CN.png', 'en.png', 'BG.png', 'zh-CN.jpg', 'en.jpg'];
  for (const name of preferred) {
    if (files.includes(name)) return name;
  }
  return files[0];
}

function handlePreviewHtml(req, res) {
  const urlObj = new URL(req.url, `http://localhost:${PORT}`);
  const templateKey = urlObj.searchParams.get('template');
  if (!templateKey) {
    res.writeHead(400);
    return res.end('缺少 template 参数');
  }

  try {
    const templates = JSON.parse(fs.readFileSync(path.join(__dirname, 'templates.json'), 'utf-8'));
    const tpl = templates.templates.find(t => t.key === templateKey);
    if (!tpl) {
      res.writeHead(404);
      return res.end('模板不存在');
    }

    const posterFile = path.join(__dirname, tpl.file);
    let html = fs.readFileSync(posterFile, 'utf-8');

    const bgDirName = PREVIEW_BG_DIRS[templateKey];
    const bgFile = bgDirName ? pickBgFile(path.join(__dirname, bgDirName)) : null;

    // Inject base tag so all relative asset paths (team icons, logos, images) resolve correctly
    const baseTag = `<base href="http://localhost:${PORT}/">`;
    const bgInject = bgFile
      ? `<script>window.BG_PATH = "http://localhost:${PORT}/${bgDirName}/${bgFile}";</script>`
      : '';

    html = html.replace('<head>', `<head>\n  ${baseTag}`);
    if (bgInject) {
      html = html.replace('</head>', bgInject + '\n</head>');
    }

    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(html);
  } catch (err) {
    res.writeHead(500);
    res.end(`预览失败：${err.message}`);
  }
}

function handleDownloadZip(req, res) {
  const urlObj = new URL(req.url, `http://${req.headers.host || `localhost:${PORT}`}`);
  const relPath = urlObj.searchParams.get('path');
  if (!relPath) {
    res.writeHead(400);
    return res.end('缺少 path 参数');
  }
  const absPath = path.resolve(__dirname, relPath);
  if (!absPath.startsWith(OUTPUT_ROOT + path.sep) || !absPath.toLowerCase().endsWith('.zip')) {
    res.writeHead(403);
    return res.end('forbidden');
  }
  if (!fs.existsSync(absPath)) {
    res.writeHead(404);
    return res.end('zip 不存在');
  }
  const filename = path.basename(absPath);
  res.writeHead(200, {
    'Content-Type': 'application/zip',
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
    'Content-Length': fs.statSync(absPath).size,
  });
  fs.createReadStream(absPath).pipe(res);
}

function handleUsage(req, res) {
  const log = readUsageLog();
  const today = todayKey();
  const todayEntries = log.filter(e => (e.ts || '').startsWith(today));
  const startsToday = todayEntries.filter(e => e.event === 'start').length;
  const endsToday = todayEntries.filter(e => e.event === 'end');
  const successToday = endsToday.filter(e => e.status === 'success').length;
  const failToday = endsToday.filter(e => e.status === 'fail').length;
  const recent = log.filter(e => e.event === 'end').slice(-30).reverse();
  const oneWeekAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const last7Days = log.filter(e => e.event === 'start' && new Date(e.ts).getTime() >= oneWeekAgo).length;
  res.writeHead(200, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({
    todayCount: startsToday,
    todaySuccess: successToday,
    todayFail: failToday,
    quota: DAILY_QUOTA,
    quotaRemaining: Math.max(0, DAILY_QUOTA - startsToday),
    last7Days,
    recent,
  }));
}

const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }

  // Polymarket 代理：/api/polymarket/*
  if (req.url.startsWith('/api/polymarket/')) {
    const apiPath = req.url.replace('/api/polymarket', '');
    return proxyPolymarket(req, res, apiPath);
  }

  // 返回模板列表
  if (req.url === '/api/templates' && req.method === 'GET') {
    const data = fs.readFileSync(path.join(__dirname, 'templates.json'), 'utf-8');
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(data);
  }

  // SSE 生成流
  if (req.url.startsWith('/api/generate-stream') && req.method === 'GET') {
    return handleGenerateStream(req, res);
  }

  // 模板预览 HTML
  if (req.url.startsWith('/api/preview-html') && req.method === 'GET') {
    return handlePreviewHtml(req, res);
  }

  // 下载 zip
  if (req.url.startsWith('/api/download-zip') && req.method === 'GET') {
    return handleDownloadZip(req, res);
  }

  // 使用统计
  if (req.url === '/api/usage' && req.method === 'GET') {
    return handleUsage(req, res);
  }

  // 静态文件服务
  let filePath = path.join(__dirname, req.url === '/' ? '/index.html' : decodeURIComponent(req.url.split('?')[0]));
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      return res.end('Not found');
    }
    const ext = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`\n🎨 海报生成器已启动：http://localhost:${PORT}（监听 0.0.0.0，全网卡可访问）`);
  console.log(`日额度：${DAILY_QUOTA} 次/天（设置环境变量 DAILY_QUOTA 可调整）`);
  console.log(`使用日志：${path.relative(process.cwd(), USAGE_LOG)}\n`);
});
