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
};

let isGenerating = false;

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

  if (!templateKey) {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ error: '缺少 template 参数' }));
  }

  if (isGenerating) {
    res.writeHead(409, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ error: '正在生成中，请稍候' }));
  }

  isGenerating = true;

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

  sendEvent('message', `开始生成模板：${templateKey}`);

  const child = spawn('node', ['generate.js', '--template', templateKey], {
    cwd: __dirname,
    env: { ...process.env, FORCE_COLOR: '0' },
  });

  const onLine = (chunk) => {
    const lines = chunk.toString().split('\n');
    lines.forEach(line => {
      if (line.trim()) sendEvent('message', line.replace(/\r/g, ''));
    });
  };

  child.stdout.on('data', onLine);
  child.stderr.on('data', onLine);

  child.on('close', (code) => {
    isGenerating = false;
    if (code === 0) {
      sendEvent('done', '生成完成');
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
    const bgInject = bgFile
      ? `<script>window.BG_PATH = "http://localhost:${PORT}/${bgDirName}/${bgFile}";</script>`
      : '';

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

function handleOpenOutput(req, res, body) {
  let payload;
  try {
    payload = JSON.parse(body);
  } catch {
    res.writeHead(400);
    return res.end(JSON.stringify({ error: '无效 JSON' }));
  }
  const dir = payload.dir;
  if (!dir) {
    res.writeHead(400);
    return res.end(JSON.stringify({ error: '缺少 dir 参数' }));
  }
  const absDir = path.resolve(__dirname, dir);
  spawn('open', [absDir]);
  res.writeHead(200, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ ok: true }));
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

  // 打开输出文件夹
  if (req.url === '/api/open-output' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => handleOpenOutput(req, res, body));
    return;
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

server.listen(PORT, () => {
  console.log(`\n🎨 海报生成器已启动：http://localhost:${PORT}`);
  console.log(`Polymarket 代理：http://localhost:${PORT}/api/polymarket/events/slug/f1-drivers-champion\n`);
});
