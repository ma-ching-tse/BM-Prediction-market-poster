// 轻量 Polymarket 代理 + 静态文件服务器（预览用）
// 用法：node proxy.js
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

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

const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');

  // Polymarket 代理：/api/polymarket/*
  if (req.url.startsWith('/api/polymarket/')) {
    const apiPath = req.url.replace('/api/polymarket', '');
    return proxyPolymarket(req, res, apiPath);
  }

  // 静态文件服务
  let filePath = path.join(__dirname, req.url === '/' ? '/poster.f1driver.html' : decodeURIComponent(req.url.split('?')[0]));
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
  console.log(`预览代理服务器运行在 http://localhost:${PORT}`);
  console.log(`Polymarket 代理：http://localhost:${PORT}/api/polymarket/events/slug/f1-drivers-champion`);
});
