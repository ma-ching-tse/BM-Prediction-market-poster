const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const MAX_SIZE_KB = 280;
const POLYMARKET_BASE_URL = 'https://gamma-api.polymarket.com';

// ── sharp 方案：缩放到指定尺寸，并对 JPEG 质量进行二分，逼近 maxKB 上限 ──
async function compressToJpegWithSharp(inputPath, outputPath, maxKB, outputSize) {
  const sharp = require('sharp');
  const outW = Number(outputSize.width);
  const outH = Number(outputSize.height);

  const render = async (quality) => {
    return await sharp(inputPath)
      .resize(outW, outH, { fit: 'cover', position: 'centre' })
      .jpeg({ quality, chromaSubsampling: '4:4:4', mozjpeg: true })
      .toBuffer();
  };

  // 先看最高质量能不能直接过（常见情况可避免二分）
  let bestBuf = await render(95);
  let bestQuality = 95;
  let bestSizeKB = Math.round(bestBuf.length / 1024);

  if (bestSizeKB > maxKB) {
    // 二分 [40, 95]，找出 ≤ maxKB 的最高质量
    let lo = 40;
    let hi = 95;
    while (lo <= hi) {
      const mid = Math.round((lo + hi) / 2);
      const buf = await render(mid);
      const kb = Math.round(buf.length / 1024);
      if (kb <= maxKB) {
        bestBuf = buf;
        bestQuality = mid;
        bestSizeKB = kb;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
  }

  fs.writeFileSync(outputPath, bestBuf);

  return {
    sizeKB: bestSizeKB,
    profile: bestSizeKB <= maxKB ? `within-limit/q${bestQuality}` : `best-effort/q${bestQuality}`
  };
}

const BASE_DIR = __dirname;
const TEAMS_CSV    = path.join(BASE_DIR, 'teams.csv');
const FOOTBALL_TEAMS_CSV = path.join(BASE_DIR, 'football_teams.csv');
const F1_TEAMS_CSV = path.join(BASE_DIR, 'teams_f1.csv');
const F1_ICON_DIR  = path.join(BASE_DIR, 'assets', 'logos', 'F1 车队');
const F1_DRIVER_CSV     = path.join(BASE_DIR, 'drivers_f1.csv');
const F1_DRIVER_ICON_DIR = path.join(BASE_DIR, 'assets', 'logos', 'F1 Driver');
const BG_DIR       = path.join(BASE_DIR, 'backgrounds-NBA');
const WORLD_CUP_BG_DIR = path.join(BASE_DIR, 'backgrounds- football');
const GLOBAL_BG_DIR = path.join(BASE_DIR, 'backgrounds-global');
// 统一固定输出到当前项目目录，避免写到上级目录造成混淆。
const OUTPUT_DIR = path.join(BASE_DIR, 'output');
const POSTER_COPY_CONFIG = path.join(BASE_DIR, 'poster.copy.json');

// lark.config.json 可能在主目录（worktree 上级），逐级向上查找
const LARK_CONFIG = (() => {
  const candidates = [
    path.join(BASE_DIR, 'lark.config.json'),
    path.join(BASE_DIR, '..', '..', '..', 'lark.config.json'),  // worktree: /.claude/worktrees/NAME/
  ];
  return candidates.find(p => fs.existsSync(p)) ?? candidates[0];
})();
const TEMPLATE_CONFIGS = {
  classic: {
    aliases: ['classic', 'default', '标准', '默认'],
    file: path.join(BASE_DIR, 'poster.html'),
    horizontalFile: path.join(BASE_DIR, 'poster.nba-horizontal.html'),
    horizontalBgDir: path.join(BASE_DIR, 'backgrounds-NBA-horizontal'),
    outputPrefix: 'NBA',
    outputSubDir: 'NBA'
  },
  comprehensive: {
    aliases: ['comprehensive', 'event', '综合事件', '综合事件模版'],
    file: path.join(BASE_DIR, 'poster.comprehensive-event.html'),
    outputPrefix: '综合事件',
    outputSubDir: '综合事件'
  },
  worldcup: {
    aliases: ['worldcup', 'world-cup', 'world cup', '世界杯', '世界杯模版'],
    file: path.join(BASE_DIR, 'poster.world-cup.html'),
    outputPrefix: '世界杯',
    outputSubDir: '世界杯',
    bgDir: WORLD_CUP_BG_DIR
  },
  football: {
    aliases: ['football', 'soccer', 'football-soccer', '足球', '足球赛事'],
    file: path.join(BASE_DIR, 'poster.football-soccer.html'),
    outputPrefix: '足球赛事',
    outputSubDir: '足球赛事',
    bgDir: WORLD_CUP_BG_DIR,
    teamsCsv: FOOTBALL_TEAMS_CSV,
    larkSheet: 'Fleir2'
  },
  coinprice: {
    aliases: ['coinprice', 'coin-price', 'coin', '币价预测', '币价'],
    file: path.join(BASE_DIR, 'poster.coin-price.html'),
    outputPrefix: '币价预测',
    outputSubDir: '币价预测',
    bgDir: GLOBAL_BG_DIR,
    logosDir: path.join(BASE_DIR, 'assets', 'logos', '币价预测'),
    larkSheet: 'dxcuKC'
  },
  global: {
    aliases: ['global', 'global-prediction-market', '全球预测市场'],
    file: path.join(BASE_DIR, 'poster.global-prediction-market.html'),
    outputPrefix: '全球预测市场',
    outputSubDir: '全球预测市场',
    bgDir: GLOBAL_BG_DIR,
    logosDir: path.join(BASE_DIR, 'assets', 'logos', '全球预测市场'),
    larkSheet: 'ZbFFnr'
  },
  f1: {
    aliases: ['f1', 'f1赛车', 'F1', 'formula1', 'formula-1'],
    file: path.join(BASE_DIR, 'poster.f1.html'),
    outputPrefix: 'F1',
    outputSubDir: 'F1 车队'
  },
  f1driver: {
    aliases: ['f1driver', 'f1-driver', 'f1车手', 'f1 driver', 'driver'],
    file: path.join(BASE_DIR, 'poster.f1driver.html'),
    outputPrefix: 'F1Driver',
    outputSubDir: 'F1 Driver'
  }
};

function resolveTemplateConfig(inputKey = 'classic') {
  const normalized = String(inputKey ?? '').trim().toLowerCase();
  for (const [key, config] of Object.entries(TEMPLATE_CONFIGS)) {
    if (key === normalized) return { key, config };
    if (config.aliases.some(alias => String(alias).toLowerCase() === normalized)) {
      return { key, config };
    }
  }

  const supported = Object.keys(TEMPLATE_CONFIGS).join(', ');
  throw new Error(`不支持的模板：${inputKey}。可用模板：${supported}`);
}

function parseCliOptions(argv = process.argv.slice(2)) {
  let templateInput = 'classic';

  for (let i = 0; i < argv.length; i++) {
    const arg = String(argv[i] ?? '');
    if (!arg) continue;

    if (arg === '--help' || arg === '-h') {
      console.log([
        '用法：node generate.js [--template <模板名>]',
        '可用模板：',
        '  classic        标准赛事卡片模板（默认）',
        '  comprehensive  综合事件模版',
        '  worldcup       世界杯模版',
        '  football       足球赛事模版',
        '  coinprice      币价预测模版',
        '  global         全球预测市场模版',
        '  f1             F1车队海报模版',
        '  f1driver       F1车手海报模版'
      ].join('\n'));
      process.exit(0);
    }

    if (arg === '--template') {
      const nextArg = String(argv[i + 1] ?? '').trim();
      if (!nextArg) {
        throw new Error('参数 --template 需要传入模板名');
      }
      templateInput = nextArg;
      i++;
      continue;
    }

    if (arg.startsWith('--template=')) {
      templateInput = arg.slice('--template='.length).trim();
      if (!templateInput) {
        throw new Error('参数 --template= 需要传入模板名');
      }
      continue;
    }
  }

  const { key, config } = resolveTemplateConfig(templateInput);
  if (!fs.existsSync(config.file)) {
    throw new Error(`模板文件不存在：${config.file}`);
  }

  return { templateKey: key, templateConfig: config };
}

function loadPosterCopyConfig() {
  if (!fs.existsSync(POSTER_COPY_CONFIG)) {
    throw new Error(`未找到海报文案配置文件：${POSTER_COPY_CONFIG}`);
  }

  try {
    return JSON.parse(fs.readFileSync(POSTER_COPY_CONFIG, 'utf8'));
  } catch (err) {
    throw new Error(`海报文案配置不是合法 JSON：${err.message}`);
  }
}

// ── 生成过程告警去重 ─────────────────────────────────────
const warningSet = new Set();
function warnOnce(key, message) {
  if (warningSet.has(key)) return;
  warningSet.add(key);
  console.warn(`⚠️  ${message}`);
}

function loadLarkConfig() {
  if (!fs.existsSync(LARK_CONFIG)) {
    throw new Error(`未找到配置文件：${LARK_CONFIG}`);
  }

  let config;
  try {
    config = JSON.parse(fs.readFileSync(LARK_CONFIG, 'utf8'));
  } catch (err) {
    throw new Error(`配置文件不是合法 JSON：${err.message}`);
  }

  const requiredFields = ['appId', 'appSecret'];
  for (const field of requiredFields) {
    if (!String(config[field] ?? '').trim()) {
      throw new Error(`配置文件缺少必填字段：${field}`);
    }
  }

  const hasSpreadsheetToken = String(config.spreadsheetToken ?? '').trim();
  const hasSpreadsheetUrl = String(config.spreadsheetUrl ?? '').trim();
  const hasWikiToken = String(config.wikiToken ?? '').trim();
  const hasWikiUrl = String(config.wikiUrl ?? '').trim();
  if (!hasSpreadsheetToken && !hasSpreadsheetUrl && !hasWikiToken && !hasWikiUrl) {
    throw new Error('配置文件至少需要提供 spreadsheetToken、spreadsheetUrl、wikiToken、wikiUrl 其中之一');
  }

  return {
    appId: String(config.appId).trim(),
    appSecret: String(config.appSecret).trim(),
    spreadsheetToken: String(config.spreadsheetToken ?? '').trim(),
    spreadsheetUrl: String(config.spreadsheetUrl ?? '').trim(),
    wikiToken: String(config.wikiToken ?? '').trim(),
    wikiUrl: String(config.wikiUrl ?? '').trim(),
    sheetId: String(config.sheetId ?? '').trim(),
    range: String(config.range ?? 'A1:E500').trim(),
    comprehensiveSheetId: String(config.comprehensiveSheetId ?? '').trim(),
    worldCupSheetId: String(config.worldCupSheetId ?? '').trim(),
    f1SheetId: String(config.f1SheetId ?? '').trim(),
    f1SpreadsheetToken: String(config.f1SpreadsheetToken ?? '').trim(),
    f1DriverSheetId: String(config.f1DriverSheetId ?? '').trim(),
    f1DriverSpreadsheetToken: String(config.f1DriverSpreadsheetToken ?? '').trim(),
    sourceLang: String(config.sourceLang ?? 'zh-CN').trim(),
    anthropicApiKey: String(config.anthropicApiKey ?? '').trim()
  };
}

async function getLarkTenantAccessToken(config) {
  const res = await fetch('https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({
      app_id: config.appId,
      app_secret: config.appSecret
    })
  });

  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 鉴权失败：${data.msg || res.status}`);
  }
  return data.tenant_access_token;
}

function normalizeHeader(value) {
  return String(value ?? '').trim().toLowerCase();
}

function columnLabelToIndex(label) {
  let result = 0;
  const upper = String(label ?? '').trim().toUpperCase();
  for (const ch of upper) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    result = result * 26 + (code - 64);
  }
  return result - 1;
}

function indexToColumnLabel(index) {
  let num = Number(index) + 1;
  if (!Number.isInteger(num) || num <= 0) {
    throw new Error(`非法列索引：${index}`);
  }
  let result = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

function parseRangeStart(range) {
  const raw = String(range ?? '').trim();
  const rangePart = raw.includes('!') ? raw.split('!').pop() : raw;
  const startCell = String(rangePart).split(':')[0]?.trim() ?? '';
  const match = startCell.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return { row: 1, col: 0 };
  return {
    row: Number(match[2]),
    col: columnLabelToIndex(match[1])
  };
}

function extractTokenFromLarkUrl(rawUrl, expectedType) {
  if (!rawUrl) return '';

  let parsed;
  try {
    parsed = new URL(rawUrl);
  } catch {
    return '';
  }

  const parts = parsed.pathname.split('/').filter(Boolean);
  const markerIndex = parts.findIndex(part => part === expectedType);
  if (markerIndex === -1) return '';
  return parts[markerIndex + 1] ?? '';
}

async function resolveWikiNode(config, accessToken) {
  const wikiToken = config.wikiToken || extractTokenFromLarkUrl(config.wikiUrl, 'wiki');
  if (!wikiToken) {
    throw new Error('未能从 wiki 链接中解析出 wikiToken');
  }

  const url = new URL('https://open.larksuite.com/open-apis/wiki/v2/spaces/get_node');
  url.searchParams.set('token', wikiToken);

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });

  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark Wiki 节点解析失败：${data.msg || res.status}`);
  }

  const node = data.data?.node ?? data.data ?? {};
  const objType = node.obj_type ?? node.objType ?? '';
  const objToken = node.obj_token ?? node.objToken ?? '';

  if (!objToken) {
    throw new Error('Lark Wiki 节点解析失败：未返回 obj_token');
  }

  if (String(objType).toLowerCase().includes('bitable') || String(objType) === '8') {
    throw new Error('当前 wiki 链接指向的是 Lark Base，而不是普通电子表格，请改用 sheets 链接或切回 Base 方案');
  }

  return objToken;
}

async function resolveSpreadsheetToken(config, accessToken) {
  if (config.spreadsheetToken) return config.spreadsheetToken;

  const spreadsheetTokenFromUrl = extractTokenFromLarkUrl(config.spreadsheetUrl, 'sheets');
  if (spreadsheetTokenFromUrl) return spreadsheetTokenFromUrl;

  return resolveWikiNode(config, accessToken);
}

function pickFirstSheetId(metainfo) {
  const sheets = metainfo?.sheets;
  if (Array.isArray(sheets) && sheets.length > 0) {
    const first = sheets[0];
    return first.sheetId || first.sheet_id || first?.properties?.sheetId || first?.properties?.sheet_id || '';
  }

  if (sheets && typeof sheets === 'object') {
    for (const value of Object.values(sheets)) {
      const sheetId = value?.sheetId || value?.sheet_id || value?.properties?.sheetId || value?.properties?.sheet_id || '';
      if (sheetId) return sheetId;
    }
  }

  return '';
}

function listSheetInfos(metainfo) {
  const sheets = metainfo?.sheets;
  const infos = [];

  if (Array.isArray(sheets)) {
    for (const item of sheets) {
      const sheetId = item?.sheetId || item?.sheet_id || item?.properties?.sheetId || item?.properties?.sheet_id || '';
      const title = item?.title || item?.sheetName || item?.sheet_name || item?.properties?.title || '';
      if (sheetId) infos.push({ sheetId: String(sheetId), title: String(title || '').trim() });
    }
    return infos;
  }

  if (sheets && typeof sheets === 'object') {
    for (const value of Object.values(sheets)) {
      const sheetId = value?.sheetId || value?.sheet_id || value?.properties?.sheetId || value?.properties?.sheet_id || '';
      const title = value?.title || value?.sheetName || value?.sheet_name || value?.properties?.title || '';
      if (sheetId) infos.push({ sheetId: String(sheetId), title: String(title || '').trim() });
    }
  }

  return infos;
}

async function fetchLarkSheetMetainfo(accessToken, spreadsheetToken) {
  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/metainfo`, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });

  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 工作表信息读取失败：${data.msg || res.status}`);
  }

  return data.data;
}

async function resolveSheetId(config, accessToken, spreadsheetToken) {
  if (config.sheetId) return config.sheetId;

  const metainfo = await fetchLarkSheetMetainfo(accessToken, spreadsheetToken);
  const sheetId = pickFirstSheetId(metainfo);
  if (!sheetId) {
    throw new Error('未能自动解析出 sheetId，请在配置文件中手动填写 sheetId');
  }

  return sheetId;
}

async function resolveSheetName(config, accessToken, spreadsheetToken, sheetId) {
  if (config.sheetName) return String(config.sheetName).trim();

  const metainfo = await fetchLarkSheetMetainfo(accessToken, spreadsheetToken);
  const matched = listSheetInfos(metainfo).find(item => item.sheetId === String(sheetId));
  if (matched?.title) return matched.title;

  return String(sheetId);
}

function cellToText(cell) {
  if (cell == null) return '';
  if (typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean') {
    return String(cell).trim();
  }
  if (typeof cell === 'object' && typeof cell.text === 'string') {
    return cell.text.trim();
  }
  return String(cell).trim();
}

function findHeaderIndex(headers, candidates) {
  const normalizedHeaders = headers.map(normalizeHeader);
  for (const candidate of candidates) {
    const index = normalizedHeaders.indexOf(normalizeHeader(candidate));
    if (index !== -1) return index;
  }
  return -1;
}

function getCellByHeaderAliases(headers, row, candidates, fallback = '') {
  const index = findHeaderIndex(headers, candidates);
  if (index === -1) return fallback;
  return String(row[index] ?? '').trim() || fallback;
}

function parseGamesFromSheetRows(values, options = {}) {
  if (!Array.isArray(values) || values.length < 2) {
    throw new Error('Lark 表格里没有可用数据，至少需要 1 行表头和 1 行内容');
  }
  const rangeStartRow = Number(options.rangeStartRow ?? 1);

  const headers = values[0].map(cellToText);
  const requiredHeaderAliases = {
    date: ['date', '日期', '比赛日期'],
    home_team: ['home_team', 'home team', '主队', '主队id', '主队ID'],
    away_team: ['away_team', 'away team', '客队', '客队id', '客队ID']
  };
  const optionalHeaderAliases = {
    home_win: ['home_win', 'home win', '主队胜率', '主胜率'],
    away_win: ['away_win', 'away win', '客队胜率', '客胜率'],
    polymarket_url: [
      'polymarket_url',
      'polymarket url',
      'polymarket_link',
      'polymarket link',
      'market_url',
      'market url',
      'match_link',
      'match link',
      '链接',
      'polymarket链接',
      'polymarket网址'
    ],
    polymarket_slug: ['polymarket_slug', 'polymarket slug', 'market_slug', 'market slug', 'slug', '赔率slug'],
    home_outcome: ['home_outcome', 'home outcome', '主队outcome', '主队 outcome'],
    away_outcome: ['away_outcome', 'away outcome', '客队outcome', '客队 outcome']
  };

  const indexes = {};
  for (const [key, aliases] of Object.entries(requiredHeaderAliases)) {
    const index = findHeaderIndex(headers, aliases);
    if (index === -1) {
      throw new Error(`Lark 表格缺少字段：${key}`);
    }
    indexes[key] = index;
  }
  for (const [key, aliases] of Object.entries(optionalHeaderAliases)) {
    const index = findHeaderIndex(headers, aliases);
    if (index !== -1) {
      indexes[key] = index;
    }
  }

  const rows = values.slice(1)
    .map((row, rowOffset) => ({ row: row.map(cellToText), rowOffset }))
    .filter(item => item.row.some(Boolean))
    .map(item => ({
      date: item.row[indexes.date] ?? '',
      home_team: item.row[indexes.home_team] ?? '',
      away_team: item.row[indexes.away_team] ?? '',
      home_win: String(item.row[indexes.home_win] ?? '').replace('%', '').trim(),
      away_win: String(item.row[indexes.away_win] ?? '').replace('%', '').trim(),
      polymarket_url: String(item.row[indexes.polymarket_url] ?? '').trim(),
      polymarket_slug: String(item.row[indexes.polymarket_slug] ?? '').trim(),
      home_outcome: String(item.row[indexes.home_outcome] ?? '').trim(),
      away_outcome: String(item.row[indexes.away_outcome] ?? '').trim(),
      __sheetRowNumber: rangeStartRow + 1 + item.rowOffset
    }));

  return { rows, indexes };
}

// ── 从 Lark 普通电子表格读取比赛数据 ───────────────────
async function fetchGamesFromLarkSheets() {
  const config = loadLarkConfig();
  const accessToken = await getLarkTenantAccessToken(config);
  const spreadsheetToken = await resolveSpreadsheetToken(config, accessToken);
  const sheetId = await resolveSheetId(config, accessToken, spreadsheetToken);
  const sheetName = await resolveSheetName(config, accessToken, spreadsheetToken, sheetId);
  const rangeStart = parseRangeStart(config.range);
  const range = `${sheetId}!${config.range}`;
  const url = new URL(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encodeURIComponent(range)}`);
  url.searchParams.set('valueRenderOption', 'ToString');
  url.searchParams.set('dateTimeRenderOption', 'FormattedString');

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });

  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  const parsed = parseGamesFromSheetRows(values, { rangeStartRow: rangeStart.row });
  return {
    rows: parsed.rows,
    headerIndexes: parsed.indexes,
    larkContext: {
      accessToken,
      spreadsheetToken,
      sheetId,
      sheetName,
      rangeStart
    }
  };
}

function formatPercentForSheet(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return '';
  const rounded = round2(numeric);
  return Number.isInteger(rounded) ? String(rounded) : String(rounded);
}

async function updateLarkCellValue({ accessToken, spreadsheetToken, sheetId, a1Cell, value }) {
  const sheetTarget = String(sheetId ?? '').trim();
  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({
      valueRange: {
        range: `${sheetTarget}!${a1Cell}:${a1Cell}`,
        values: [[value]]
      }
    })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 回填失败（${a1Cell}）：${data.msg || res.status}`);
  }
}

async function writeBackWinRatesToLark(rows, headerIndexes, larkContext) {
  const homeWinCol = headerIndexes.home_win;
  const awayWinCol = headerIndexes.away_win;
  const polymarketUrlCol = headerIndexes.polymarket_url;
  const polymarketSlugCol = headerIndexes.polymarket_slug;

  if (!Number.isInteger(homeWinCol) || !Number.isInteger(awayWinCol)) {
    throw new Error('Lark 表格缺少 home_win / away_win 列，无法回填赔率');
  }

  const updatedRows = rows.filter(row => row.__autoFilledFromPolymarket);
  for (const row of updatedRows) {
    const homeColA1 = indexToColumnLabel(larkContext.rangeStart.col + homeWinCol);
    const awayColA1 = indexToColumnLabel(larkContext.rangeStart.col + awayWinCol);
    const rowNumber = row.__sheetRowNumber;

    await updateLarkCellValue({
      accessToken: larkContext.accessToken,
      spreadsheetToken: larkContext.spreadsheetToken,
      sheetId: larkContext.sheetId,
      a1Cell: `${homeColA1}${rowNumber}`,
      value: formatPercentForSheet(row.home_win)
    });
    await updateLarkCellValue({
      accessToken: larkContext.accessToken,
      spreadsheetToken: larkContext.spreadsheetToken,
      sheetId: larkContext.sheetId,
      a1Cell: `${awayColA1}${rowNumber}`,
      value: formatPercentForSheet(row.away_win)
    });

    if (Number.isInteger(polymarketUrlCol) && row.__shouldWritePolymarketUrlBack && String(row.__resolvedPolymarketUrl ?? '').trim()) {
      const polymarketUrlColA1 = indexToColumnLabel(larkContext.rangeStart.col + polymarketUrlCol);
      await updateLarkCellValue({
        accessToken: larkContext.accessToken,
        spreadsheetToken: larkContext.spreadsheetToken,
        sheetId: larkContext.sheetId,
        a1Cell: `${polymarketUrlColA1}${rowNumber}`,
        value: row.__resolvedPolymarketUrl
      });
    }

    if (Number.isInteger(polymarketSlugCol) && row.__shouldWritePolymarketSlugBack && String(row.__resolvedPolymarketSlug ?? '').trim()) {
      const polymarketSlugColA1 = indexToColumnLabel(larkContext.rangeStart.col + polymarketSlugCol);
      await updateLarkCellValue({
        accessToken: larkContext.accessToken,
        spreadsheetToken: larkContext.spreadsheetToken,
        sheetId: larkContext.sheetId,
        a1Cell: `${polymarketSlugColA1}${rowNumber}`,
        value: row.__resolvedPolymarketSlug
      });
    }
  }

  return updatedRows.length;
}

// ── CSV 解析（teams.csv 仍用本地文件）────────────────────
function parseCSV(filePath) {
  const lines = fs.readFileSync(filePath, 'utf8').trim().split('\n');
  const headers = lines[0].split(',').map(h => h.trim());
  return lines.slice(1).map(line => {
    const values = line.split(',').map(v => v.trim());
    return Object.fromEntries(headers.map((h, i) => [h, values[i] ?? '']));
  });
}

// ── 读取球队翻译表，转成 { teamId: { 'zh-CN': '...', en: '...', ... } } ──
function loadTeams(filePath) {
  const rows = parseCSV(filePath);
  const map = {};
  for (const row of rows) {
    map[row.id] = row;
  }
  return map;
}

// ── Logo 文件名映射（用 logo 列找 NBA_icon 里对应的图片）──
function findLogoPath(teamId, teamsMap) {
  const logoName = teamsMap[teamId]?.['logo'];
  if (!logoName) return null;
  const normalizedLogo = String(logoName).trim();
  const logoPath = normalizedLogo.includes('/')
    ? path.join(BASE_DIR, normalizedLogo)
    : path.join(BASE_DIR, 'NBA_icon', `${normalizedLogo}.png`);
  return fs.existsSync(logoPath) ? logoPath : null;
}

// ── F1：Logo 文件路径（assets/logos/F1 车队/constructors_{logo}.png）──
function findF1LogoPath(teamId, f1TeamsMap) {
  const logoName = f1TeamsMap[teamId]?.['logo'];
  if (!logoName) return null;
  const logoPath = path.join(F1_ICON_DIR, `constructors_${logoName}.png`);
  return fs.existsSync(logoPath) ? logoPath : null;
}

// ── F1：Polymarket 车队/车手匹配 ─────────────────────────
function resolveF1TeamByOutcome(outcomeText, f1TeamsMap) {
  const normalized = normalizeOutcomeToken(outcomeText);
  if (!normalized) return null;
  for (const [id, team] of Object.entries(f1TeamsMap)) {
    const aliases = [id, team.en, team['zh-CN'], team['zh-TW'], team.ja, team.logo]
      .map(v => normalizeOutcomeToken(v)).filter(Boolean);
    if (aliases.some(a => a === normalized || normalized.includes(a) || a.includes(normalized))) return id;
  }
  return null;
}

// 从问题文本中提取主语（如 "Will Ferrari be..." → "Ferrari"）
function extractTeamNameFromQuestion(question) {
  const m = String(question ?? '').match(/^Will\s+(.+?)\s+be\s+/i);
  return m ? m[1].trim() : null;
}

// 适用于多选市场（如分站赛冠军，outcomes = [Driver1, Driver2, ...]）
function buildF1TeamsFromPolymarketMarket(market, f1TeamsMap) {
  const slug = String(market?.slug ?? '').trim() || 'unknown';
  const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(s => String(s).trim());
  const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', slug);
  if (outcomes.length !== prices.length) {
    throw new Error(`Polymarket 市场（${slug}）outcomes 与 outcomePrices 数量不匹配`);
  }
  return outcomes.map((outcome, i) => {
    const percent = toPercentProbability(prices[i]);
    if (!Number.isFinite(percent)) return null;
    const teamId = resolveF1TeamByOutcome(outcome, f1TeamsMap);
    return {
      teamId: teamId ?? `__raw__${outcome}`,
      _rawName: outcome,
      percent: Math.max(0, Math.min(100, Math.round(percent)))
    };
  }).filter(Boolean).sort((a, b) => b.percent - a.percent).slice(0, 3);
}

// 适用于多 Yes/No 子市场的 event（如总冠军，每个车队一个 Yes/No 市场）
function buildF1TeamsFromPolymarketYesNoEvent(event, f1TeamsMap) {
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  const teamProbs = [];

  for (const market of markets) {
    const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', market?.slug ?? '')
      .map(s => String(s).trim());
    const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', market?.slug ?? '');
    const yesIdx = outcomes.findIndex(o => o.toLowerCase() === 'yes');
    if (yesIdx === -1) continue; // 不是 Yes/No 市场，跳过

    const percent = toPercentProbability(prices[yesIdx]);
    if (!Number.isFinite(percent)) continue;

    // 从 question 提取车队名，如 "Will Ferrari be..." → "Ferrari"
    const rawName = extractTeamNameFromQuestion(market.question)
      ?? String(market.slug ?? '').replace(/^will-/, '').replace(/-be-.*$/, '').replace(/-/g, ' ');

    const teamId = resolveF1TeamByOutcome(rawName, f1TeamsMap);
    teamProbs.push({
      teamId: teamId ?? `__raw__${rawName}`,
      _rawName: rawName,
      percent: Math.max(0, Math.min(100, Math.round(percent)))
    });
  }

  return teamProbs.sort((a, b) => b.percent - a.percent).slice(0, 3);
}

async function fetchF1TeamsFromPolymarket(inputSlugOrUrl, f1TeamsMap) {
  const slug = extractPolymarketSlug(inputSlugOrUrl);
  if (!slug) return null;

  // 先尝试单个 market（适合分站赛多选市场）
  let market = null;
  try {
    market = await fetchPolymarketMarketBySlug(slug);
  } catch (err) {
    if (!String(err?.message ?? '').includes('HTTP 404')) throw err;
  }
  if (market) return buildF1TeamsFromPolymarketMarket(market, f1TeamsMap);

  // 再尝试 event（适合总冠军等多个 Yes/No 子市场）
  const event = await fetchPolymarketEventBySlug(slug);
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  if (markets.length === 0) throw new Error(`Polymarket 事件（${slug}）没有 markets 数据`);

  // 判断是 Yes/No 子市场模式还是多选市场模式
  const firstOutcomes = parseJsonArrayField(markets[0]?.outcomes ?? '[]', 'outcomes', '').map(s => String(s).toLowerCase().trim());
  const isYesNoEvent = firstOutcomes.includes('yes') && firstOutcomes.includes('no');

  if (isYesNoEvent) {
    return buildF1TeamsFromPolymarketYesNoEvent(event, f1TeamsMap);
  }

  // 多选市场：选 outcomes 最多的那个
  const mainMarket = markets.reduce((best, m) => {
    const blen = parseJsonArrayField(best?.outcomes ?? '[]', 'outcomes', 'best').length;
    const mlen = parseJsonArrayField(m?.outcomes ?? '[]', 'outcomes', 'cur').length;
    return mlen > blen ? m : best;
  });
  return buildF1TeamsFromPolymarketMarket(mainMarket, f1TeamsMap);
}

// ── F1：从 Lark 读取事件数据 ───────────────────────────
// 横向表头（A1:J1）：main title | sub title | footer | team_1 | percent_1 | team_2 | percent_2 | team_3 | percent_3
// 竖向配置区（A列=key, B列=value）：market_slug = <Polymarket URL>
async function fetchF1EventsFromLark(config, accessToken) {
  const sheetId = String(config.f1SheetId ?? '').trim();
  if (!sheetId) {
    throw new Error('lark.config.json 缺少 f1SheetId');
  }

  // 优先用 f1SpreadsheetToken（F1 可能在独立表格），否则沿用主表格
  const spreadsheetToken = config.f1SpreadsheetToken || config.spreadsheetToken;
  if (!spreadsheetToken) {
    throw new Error('lark.config.json 缺少 spreadsheetToken 或 f1SpreadsheetToken');
  }

  // 读取足够大的范围以覆盖竖向配置区（最多 30 行）
  const range = encodeURIComponent(`${sheetId}!A1:J30`);
  const url = `https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark F1 表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('F1 表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  // 第1行横向表头，第2行横向数据
  const headers = values[0].map(cell => String(cell ?? '').trim().toLowerCase());
  const row = values[1].map(cell => String(cell ?? '').trim());

  function col(name) {
    const idx = headers.indexOf(name);
    return idx !== -1 ? row[idx] : '';
  }

  // 扫描竖向配置区：A列=key, B列=value（如 market_slug）
  const kvAliases = { market_slug: 'polymarket_url', polymarket_url: 'polymarket_url', polymarket_slug: 'polymarket_url' };
  const kvConfig = {};
  for (const rowData of values) {
    const key = String(rowData[0] ?? '').trim().toLowerCase().replace(/\s+/g, '_');
    const val = String(rowData[1] ?? '').trim();
    if (kvAliases[key] && val) {
      kvConfig[kvAliases[key]] = val;
    }
  }

  return {
    mainTitle: col('main title'),
    subTitle: col('sub title'),
    footer: col('footer'),
    polymarket_url: kvConfig.polymarket_url ?? '',
    teams: [
      { teamId: col('team_1'), percent: Number(col('percent_1')) || 33 },
      { teamId: col('team_2'), percent: Number(col('percent_2')) || 33 },
      { teamId: col('team_3'), percent: Number(col('percent_3')) || 34 }
    ]
  };
}

// ── F1：翻译标题/副标题/footer（不翻译车队名，CSV 里已有）──
async function translateF1Titles(sourceData, targetLangs, fromLang = 'zh-CN') {
  const texts = [sourceData.mainTitle, sourceData.subTitle, sourceData.footer];
  const result = {};

  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(texts.map(t => translateOneText(t, fromLang, lang)));
    result[lang] = {
      mainTitle: translated[0],
      subTitle: translated[1],
      footer: translated[2]
    };
    console.log(' ✅');
  }

  return result;
}

// ── F1：翻译结果回填到 Lark 表格（第 3 行起，A~D 列）──
async function writeBackF1TranslationsToLark(sourceData, translationsMap, accessToken, spreadsheetToken, sheetId) {
  const langs = Object.keys(translationsMap);
  const rows = langs.map(lang => {
    const d = translationsMap[lang];
    return [
      lang,
      String(d.mainTitle ?? ''),
      String(d.subTitle ?? ''),
      String(d.footer ?? '')
    ];
  });

  const startRow = 3;
  const endRow = startRow + rows.length - 1;
  const range = `${sheetId}!A${startRow}:D${endRow}`;

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({ valueRange: { range, values: rows } })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`F1 翻译回填 Lark 失败：${data.msg || res.status}`);
  }
  return rows.length;
}

// ── F1：构建海报 payload ──────────────────────────────
function buildF1PosterPayload(sourceData, translationsMap, f1TeamsMap, lang, copyConfig) {
  const templateCopy = copyConfig?.f1 ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };

  const translated = translationsMap[lang] ?? sourceData;

  const teamsData = sourceData.teams.map(entry => {
    // Polymarket 未识别的车队/车手：直接用原始名称，不查 CSV
    if (String(entry.teamId ?? '').startsWith('__raw__')) {
      return {
        name: String(entry._rawName ?? entry.teamId.replace('__raw__', '')).trim(),
        logo: '',
        percent: entry.percent
      };
    }

    const teamInfo = f1TeamsMap[entry.teamId] ?? {};
    const teamName = teamInfo[lang] ?? teamInfo['en'] ?? entry.teamId;
    const logoPath = findF1LogoPath(entry.teamId, f1TeamsMap);

    if (!teamInfo[lang] && !teamInfo['en']) {
      warnOnce(`f1-team-missing:${entry.teamId}`, `F1车队 ID 不存在：${entry.teamId}`);
    }
    if (teamInfo[lang] && !logoPath) {
      warnOnce(`f1-logo-missing:${entry.teamId}`, `找不到 F1 Logo：${entry.teamId}`);
    }

    return {
      name: teamName,
      logo: logoPath ?? '',
      percent: entry.percent
    };
  });

  return {
    teams: teamsData,
    copy: {
      title: String(translated.mainTitle ?? '').trim(),
      subtitle: String(translated.subTitle ?? '').trim(),
      footer: String(translated.footer ?? '').trim(),
      titleFontSize: Number(mergedCopy.titleFontSize ?? 110),
      titleLineHeight: Number(mergedCopy.titleLineHeight ?? 1.2),
      titleMaxWidth: Number(mergedCopy.titleMaxWidth ?? 824),
      subtitleFontSize: Number(mergedCopy.subtitleFontSize ?? 46),
      subtitleLineHeight: Number(mergedCopy.subtitleLineHeight ?? 1.3),
      teamNameFontSize: Number(mergedCopy.teamNameFontSize ?? 52),
      headerImage: String(mergedCopy.headerImage ?? 'assets/f1_car.png'),
      brandLogo: String(defaultCopy.brandLogo ?? '')
    }
  };
}

// ── F1 车手：Photo 文件路径（assets/logos/F1 Driver/{photo}.png）──
function findF1DriverPhotoPath(driverId, driversMap) {
  const photoName = driversMap[driverId]?.['photo'];
  if (!photoName) return null;
  const photoPath = path.join(F1_DRIVER_ICON_DIR, `${photoName}.png`);
  return fs.existsSync(photoPath) ? photoPath : null;
}

// ── F1 车手：Polymarket outcome 匹配驾驶员 ──
function resolveF1DriverByOutcome(outcomeText, driversMap) {
  const normalized = normalizeOutcomeToken(outcomeText);
  if (!normalized) return null;
  for (const [id, driver] of Object.entries(driversMap)) {
    const aliases = [id, driver.en, driver['zh-CN'], driver['zh-TW'], driver.ja, driver.photo]
      .map(v => normalizeOutcomeToken(v)).filter(Boolean);
    if (aliases.some(a => a === normalized || normalized.includes(a) || a.includes(normalized))) return id;
  }
  return null;
}

// ── F1 车手：从多选市场构建排名（分站赛冠军等）──
function buildF1DriversFromPolymarketMarket(market, driversMap) {
  const slug = String(market?.slug ?? '').trim() || 'unknown';
  const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(s => String(s).trim());
  const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', slug);
  if (outcomes.length !== prices.length) {
    throw new Error(`Polymarket 市场（${slug}）outcomes 与 outcomePrices 数量不匹配`);
  }
  return outcomes.map((outcome, i) => {
    const percent = toPercentProbability(prices[i]);
    if (!Number.isFinite(percent)) return null;
    const driverId = resolveF1DriverByOutcome(outcome, driversMap);
    return {
      driverId: driverId ?? `__raw__${outcome}`,
      _rawName: outcome,
      percent: Math.max(0, Math.min(100, Math.round(percent)))
    };
  }).filter(Boolean).sort((a, b) => b.percent - a.percent).slice(0, 4);
}

// ── F1 车手：从 Yes/No event 构建排名（总冠军等）──
function buildF1DriversFromPolymarketYesNoEvent(event, driversMap) {
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  const driverProbs = [];

  for (const market of markets) {
    const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', market?.slug ?? '')
      .map(s => String(s).trim());
    const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', market?.slug ?? '');
    const yesIdx = outcomes.findIndex(o => o.toLowerCase() === 'yes');
    if (yesIdx === -1) continue;

    const percent = toPercentProbability(prices[yesIdx]);
    if (!Number.isFinite(percent)) continue;

    const rawName = extractTeamNameFromQuestion(market.question)
      ?? String(market.slug ?? '').replace(/^will-/, '').replace(/-win-.*$/, '').replace(/-be-.*$/, '').replace(/-/g, ' ');

    const driverId = resolveF1DriverByOutcome(rawName, driversMap);
    driverProbs.push({
      driverId: driverId ?? `__raw__${rawName}`,
      _rawName: rawName,
      percent: Math.max(0, Math.min(100, Math.round(percent)))
    });
  }

  return driverProbs.sort((a, b) => b.percent - a.percent).slice(0, 4);
}

async function fetchF1DriversFromPolymarket(inputSlugOrUrl, driversMap) {
  const slug = extractPolymarketSlug(inputSlugOrUrl);
  if (!slug) return null;

  // 先尝试单个 market（适合分站赛多选市场）
  let market = null;
  try {
    market = await fetchPolymarketMarketBySlug(slug);
  } catch (err) {
    if (!String(err?.message ?? '').includes('HTTP 404')) throw err;
  }
  if (market) return buildF1DriversFromPolymarketMarket(market, driversMap);

  // 再尝试 event（适合总冠军等多个 Yes/No 子市场）
  const event = await fetchPolymarketEventBySlug(slug);
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  if (markets.length === 0) throw new Error(`Polymarket 事件（${slug}）没有 markets 数据`);

  const firstOutcomes = parseJsonArrayField(markets[0]?.outcomes ?? '[]', 'outcomes', '').map(s => String(s).toLowerCase().trim());
  const isYesNoEvent = firstOutcomes.includes('yes') && firstOutcomes.includes('no');

  if (isYesNoEvent) {
    return buildF1DriversFromPolymarketYesNoEvent(event, driversMap);
  }

  const mainMarket = markets.reduce((best, m) => {
    const blen = parseJsonArrayField(best?.outcomes ?? '[]', 'outcomes', 'best').length;
    const mlen = parseJsonArrayField(m?.outcomes ?? '[]', 'outcomes', 'cur').length;
    return mlen > blen ? m : best;
  });
  return buildF1DriversFromPolymarketMarket(mainMarket, driversMap);
}

// ── F1 车手：从 Lark 读取事件数据 ──
// 表头（A1）：main title | sub title | footer | driver_1 | percent_1 | driver_2 | percent_2 | driver_3 | percent_3 | driver_4 | percent_4
// 竖向配置：market_slug = <Polymarket URL>
async function fetchF1DriverEventsFromLark(config, accessToken) {
  const sheetId = String(config.f1DriverSheetId ?? '').trim();
  if (!sheetId) {
    throw new Error('lark.config.json 缺少 f1DriverSheetId');
  }

  const spreadsheetToken = config.f1DriverSpreadsheetToken || config.f1SpreadsheetToken || config.spreadsheetToken;
  if (!spreadsheetToken) {
    throw new Error('lark.config.json 缺少 f1DriverSpreadsheetToken / spreadsheetToken');
  }

  const range = encodeURIComponent(`${sheetId}!A1:L30`);
  const url = `https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark F1车手表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('F1车手表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  const headers = values[0].map(cell => String(cell ?? '').trim().toLowerCase());
  const row = values[1].map(cell => String(cell ?? '').trim());

  function col(name) {
    const idx = headers.indexOf(name);
    return idx !== -1 ? row[idx] : '';
  }

  const kvAliases = { market_slug: 'polymarket_url', polymarket_url: 'polymarket_url', polymarket_slug: 'polymarket_url' };
  const kvConfig = {};
  for (const rowData of values) {
    const key = String(rowData[0] ?? '').trim().toLowerCase().replace(/\s+/g, '_');
    const val = String(rowData[1] ?? '').trim();
    if (kvAliases[key] && val) {
      kvConfig[kvAliases[key]] = val;
    }
  }

  return {
    mainTitle: col('main title'),
    subTitle: col('sub title'),
    footer: col('footer'),
    polymarket_url: kvConfig.polymarket_url ?? '',
    drivers: [
      { driverId: col('driver_1'), percent: Number(col('percent_1')) || 33 },
      { driverId: col('driver_2'), percent: Number(col('percent_2')) || 33 },
      { driverId: col('driver_3'), percent: Number(col('percent_3')) || 34 },
      { driverId: col('driver_4'), percent: Number(col('percent_4')) || 0 }
    ].filter(d => d.driverId)
  };
}

// ── F1 车手：翻译标题/副标题/footer ──
async function translateF1DriverTitles(sourceData, targetLangs, fromLang = 'zh-CN') {
  const texts = [sourceData.mainTitle, sourceData.subTitle, sourceData.footer];
  const result = {};

  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(texts.map(t => translateOneText(t, fromLang, lang)));
    result[lang] = {
      mainTitle: translated[0],
      subTitle: translated[1],
      footer: translated[2]
    };
    console.log(' ✅');
  }

  return result;
}

// ── F1 车手：翻译结果回填到 Lark ──
async function writeBackF1DriverTranslationsToLark(sourceData, translationsMap, accessToken, spreadsheetToken, sheetId) {
  const langs = Object.keys(translationsMap);
  const rows = langs.map(lang => {
    const d = translationsMap[lang];
    return [
      lang,
      String(d.mainTitle ?? ''),
      String(d.subTitle ?? ''),
      String(d.footer ?? '')
    ];
  });

  const startRow = 3;
  const endRow = startRow + rows.length - 1;
  const range = `${sheetId}!A${startRow}:D${endRow}`;

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({ valueRange: { range, values: rows } })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`F1车手翻译回填 Lark 失败：${data.msg || res.status}`);
  }
  return rows.length;
}

// ── F1 车手：构建海报 payload ──
function buildF1DriverPosterPayload(sourceData, translationsMap, driversMap, lang, copyConfig) {
  const templateCopy = copyConfig?.f1driver ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };

  const translated = translationsMap[lang] ?? sourceData;

  const driversData = sourceData.drivers.map(entry => {
    // Polymarket 未识别的车手：直接用原始名称
    if (String(entry.driverId ?? '').startsWith('__raw__')) {
      return {
        name: String(entry._rawName ?? entry.driverId.replace('__raw__', '')).trim(),
        photo: '',
        percent: entry.percent
      };
    }

    const driverInfo = driversMap[entry.driverId] ?? {};
    // 车手姓名：CJK 语言用本地化名，其他语言统一用英文
    const driverName = ['zh-CN', 'zh-TW', 'ja'].includes(lang)
      ? (driverInfo[lang] ?? driverInfo['en'] ?? entry.driverId)
      : (driverInfo['en'] ?? entry.driverId);
    const photoPath = findF1DriverPhotoPath(entry.driverId, driversMap);

    if (!driverInfo['en']) {
      warnOnce(`f1-driver-missing:${entry.driverId}`, `F1车手 ID 不存在：${entry.driverId}`);
    }

    return {
      name: driverName,
      photo: photoPath ?? '',
      percent: entry.percent
    };
  });

  return {
    teams: driversData,   // HTML 模板读 payload.teams 数组
    copy: {
      title: String(translated.mainTitle ?? '').trim(),
      subtitle: String(translated.subTitle ?? '').trim(),
      footer: String(translated.footer ?? '').trim(),
      titleFontSize: Number(mergedCopy.titleFontSize ?? 110),
      titleLineHeight: Number(mergedCopy.titleLineHeight ?? 1.2),
      titleMaxWidth: Number(mergedCopy.titleMaxWidth ?? 824),
      subtitleFontSize: Number(mergedCopy.subtitleFontSize ?? 46),
      subtitleLineHeight: Number(mergedCopy.subtitleLineHeight ?? 1.3),
      driverNameFontSize: Number(mergedCopy.driverNameFontSize ?? 52),
      headerImage: String(mergedCopy.headerImage ?? 'assets/f1_car.png'),
      brandLogo: String(defaultCopy.brandLogo ?? '')
    }
  };
}

// ── 构建单张海报的 games 数据（指定语种）──
function buildGamesData(gamesRows, teamsMap, lang) {
  return gamesRows.map(row => {
    const homeId = row.home_team;
    const awayId = row.away_team;

    if (!teamsMap[homeId]) {
      warnOnce(`team-missing:${homeId}`, `球队 ID 不存在：${homeId}（home_team）`);
    }
    if (!teamsMap[awayId]) {
      warnOnce(`team-missing:${awayId}`, `球队 ID 不存在：${awayId}（away_team）`);
    }

    const homeLogo = findLogoPath(homeId, teamsMap);
    const awayLogo = findLogoPath(awayId, teamsMap);

    if (teamsMap[homeId] && !homeLogo) {
      warnOnce(`logo-missing:${homeId}`, `找不到球队 Logo：${homeId}`);
    }
    if (teamsMap[awayId] && !awayLogo) {
      warnOnce(`logo-missing:${awayId}`, `找不到球队 Logo：${awayId}`);
    }

    return {
      date: row.date,
      homeTeam: {
        name:    teamsMap[homeId]?.[lang] ?? teamsMap[homeId]?.['en'] ?? homeId,
        logo:    homeLogo ?? '',
        winRate: Number(row.home_win)
      },
      awayTeam: {
        name:    teamsMap[awayId]?.[lang] ?? teamsMap[awayId]?.['en'] ?? awayId,
        logo:    awayLogo ?? '',
        winRate: Number(row.away_win)
      }
    };
  });
}

function buildClassicPosterPayload(gamesRows, teamsMap, lang, sourceData, translationsMap, copyConfig, options = {}) {
  const layout = options.layout === 'horizontal' ? 'horizontal' : 'vertical';
  const templateCopy = copyConfig?.classic ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };
  const translated = translationsMap?.[lang] ?? sourceData ?? {};

  const baseGames = buildGamesData(gamesRows, teamsMap, lang);
  const translatedMatchTexts = Array.isArray(translated.matchTexts) ? translated.matchTexts : [];
  const games = baseGames.map((game, idx) => {
    const rawReward = String(translatedMatchTexts[idx]?.reward ?? '').trim();
    // 用劣势一方的胜率（较低那个）替换 reward 文案中的第一个 NNU 占位数字
    const homePct = Number(game.homeTeam?.winRate) || 0;
    const awayPct = Number(game.awayTeam?.winRate) || 0;
    const underdogPct = Math.min(homePct, awayPct);
    const reward = rawReward.replace(/\d+(?=\s*U)/, String(underdogPct));
    return {
      ...game,
      reward,
      news: String(translatedMatchTexts[idx]?.news ?? '').trim()
    };
  });

  const copy = layout === 'horizontal'
    ? {
      title: String(translated.mainTitle ?? '').trim(),
      subtitle: String(translated.subTitle ?? '').trim(),
      footer: String(translated.footer ?? '').trim(),
      titleFontSize: Number(mergedCopy?.horizontalTitleFontSize ?? 150),
      titleLineHeight: Number(mergedCopy?.horizontalTitleLineHeight ?? 1.2),
      titleMaxWidth: Number(mergedCopy?.horizontalTitleMaxWidth ?? 1080),
      subtitleFontSize: Number(mergedCopy?.horizontalSubtitleFontSize ?? 56),
      subtitleLineHeight: Number(mergedCopy?.horizontalSubtitleLineHeight ?? 1.3)
    }
    : {
      title: String(translated.mainTitle ?? '').trim(),
      subtitle: String(translated.subTitle ?? '').trim(),
      footer: String(translated.footer ?? '').trim(),
      titleFontSize: Number(mergedCopy?.titleFontSize ?? 86),
      titleLineHeight: Number(mergedCopy?.titleLineHeight ?? 1.15),
      titleMaxWidth: Number(mergedCopy?.titleMaxWidth ?? 860),
      subtitleFontSize: Number(mergedCopy?.subtitleFontSize ?? 46),
      subtitleLineHeight: Number(mergedCopy?.subtitleLineHeight ?? 1.3)
    };

  return { games, copy };
}

function buildPosterPayload(gamesRows, teamsMap, lang, templateKey, copyConfig) {
  const templateCopy = copyConfig?.[templateKey] ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };

  const rawCards = Array.isArray(mergedCopy?.cards)
    ? mergedCopy.cards
    : (Array.isArray(defaultCopy?.cards) ? defaultCopy.cards : []);

  const cards = rawCards
    .slice(0, 3)
    .map(item => ({
      text: String(item?.text ?? '').trim(),
      image: String(item?.image ?? '').trim(),
      valueLabel: String(item?.valueLabel ?? '').trim()
    }));

  return {
    games: buildGamesData(gamesRows, teamsMap, lang),
    copy: {
      title: String(mergedCopy?.title ?? '').trim(),
      subtitle: String(mergedCopy?.subtitle ?? '').trim(),
      footer: String(mergedCopy?.footer ?? '').trim(),
      outcomeLabel: String(mergedCopy?.outcomeLabel ?? '').trim(),
      titleFontSize: Number(mergedCopy?.titleFontSize ?? 86),
      titleLineHeight: Number(mergedCopy?.titleLineHeight ?? 1.15),
      titleMaxWidth: Number(mergedCopy?.titleMaxWidth ?? 860),
      subtitleFontSize: Number(mergedCopy?.subtitleFontSize ?? 46),
      subtitleLineHeight: Number(mergedCopy?.subtitleLineHeight ?? 1.3),
      cardTextFontSize: Number(mergedCopy?.cardTextFontSize ?? 40),
      cardTextLineHeight: Number(mergedCopy?.cardTextLineHeight ?? 1.2),
      cards
    }
  };
}

// ── 胜率校验（home + away 必须等于 100）──────────────────
function validateWinRates(gamesRows) {
  const errors = [];
  for (const row of gamesRows) {
    const homeWin = Number(row.home_win);
    const awayWin = Number(row.away_win);

    if (!Number.isFinite(homeWin) || !Number.isFinite(awayWin)) {
      errors.push(`${row.date} ${row.home_team} vs ${row.away_team}：胜率不是数字`);
      continue;
    }

    const sum = homeWin + awayWin;
    // 允许极小数值误差
    if (Math.abs(sum - 100) > 0.01) {
      errors.push(`${row.date} ${row.home_team} vs ${row.away_team}：home+away=${sum}`);
    }
  }

  if (errors.length > 0) {
    const detail = errors.map(e => `- ${e}`).join('\n');
    throw new Error(`胜率校验失败（home_win + away_win 必须等于 100）：\n${detail}`);
  }
}

// ── Classic NBA：从 Lark 读取主文案 + 3 场比赛 URL ─────────────────
async function fetchClassicDataFromLark(config, accessToken, sourceLang = 'zh-CN') {
  const spreadsheetToken = await resolveSpreadsheetToken(config, accessToken);
  const sheetId = await resolveSheetId(config, accessToken, spreadsheetToken);
  const sheetName = await resolveSheetName(config, accessToken, spreadsheetToken, sheetId);
  const range = `${sheetId}!A1:Q60`;
  const url = new URL(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encodeURIComponent(range)}`);
  url.searchParams.set('valueRenderOption', 'ToString');
  url.searchParams.set('dateTimeRenderOption', 'FormattedString');

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark NBA 表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('NBA 表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  const headers = values[0].map(cell => cellToText(cell));
  const rows = values.slice(1)
    .map((row, rowOffset) => ({
      values: row.map(cell => cellToText(cell)),
      rowNumber: 2 + rowOffset
    }))
    .filter(item => item.values.some(Boolean));

  const langCol = findHeaderIndex(headers, ['lang', 'language', '语言']);
  const sourceLangNormalized = String(sourceLang || 'zh-CN').trim().toLowerCase();
  const sourceRowEntry = rows.find(item => String(item.values[langCol] ?? '').trim().toLowerCase() === sourceLangNormalized)
    ?? rows[0];

  if (!sourceRowEntry) {
    throw new Error('NBA 表格没有可用源语言行');
  }

  const sourceRow = sourceRowEntry.values;

  function getField(...aliases) {
    return getCellByHeaderAliases(headers, sourceRow, aliases);
  }

  function getMatchLink(index) {
    return getField(
      `match${index}_link`,
      `match${index}_url`,
      `match${index}_slug`,
      `match_${index}_link`,
      `match_${index}_url`,
      `match_${index}_slug`,
      `link_${index}`,
      `url_${index}`,
      `slug_${index}`
    );
  }

  const matchInputs = [];
  for (let i = 1; i <= 3; i++) {
    const link = getMatchLink(i);
    if (!link) continue;
    matchInputs.push({
      index: i,
      polymarket_url: link
    });
  }

  if (matchInputs.length === 0) {
    throw new Error('NBA 表格没有可用比赛链接（请检查 match1_link ~ match3_link 列）');
  }

  const headerIndexes = {};
  const fieldAliases = {
    title: ['title', 'main title', 'main_title', 'mian title', 'mian_title'],
    subtitle: ['subtitle', 'sub title', 'sub_title'],
    footer: ['footer', 'foot'],
    match1_home: ['match1_home', 'match_1_home'],
    match1_away: ['match1_away', 'match_1_away'],
    match1_date: ['match1_date', 'match_1_date'],
    match1_home_win: ['match1_home_win', 'match_1_home_win'],
    match1_away_win: ['match1_away_win', 'match_1_away_win'],
    match2_home: ['match2_home', 'match_2_home'],
    match2_away: ['match2_away', 'match_2_away'],
    match2_date: ['match2_date', 'match_2_date'],
    match2_home_win: ['match2_home_win', 'match_2_home_win'],
    match2_away_win: ['match2_away_win', 'match_2_away_win'],
    match3_home: ['match3_home', 'match_3_home'],
    match3_away: ['match3_away', 'match_3_away'],
    match3_date: ['match3_date', 'match_3_date'],
    match3_home_win: ['match3_home_win', 'match_3_home_win'],
    match3_away_win: ['match3_away_win', 'match_3_away_win'],
    match1_reward: ['match1_reward', 'match_1_reward'],
    match1_news: ['match1_news', 'match_1_news'],
    match2_reward: ['match2_reward', 'match_2_reward'],
    match2_news: ['match2_news', 'match_2_news'],
    match3_reward: ['match3_reward', 'match_3_reward'],
    match3_news: ['match3_news', 'match_3_news']
  };

  for (const [key, aliases] of Object.entries(fieldAliases)) {
    const index = findHeaderIndex(headers, aliases);
    if (index !== -1) headerIndexes[key] = index;
  }

  const matchTexts = [];
  for (let i = 1; i <= 3; i++) {
    matchTexts.push({
      reward: getField(`match${i}_reward`, `match_${i}_reward`),
      news: getField(`match${i}_news`, `match_${i}_news`)
    });
  }

  return {
    sourceData: {
      mainTitle: getField('title', 'main title', 'main_title', 'mian title', 'mian_title'),
      subTitle: getField('subtitle', 'sub title', 'sub_title'),
      footer: getField('footer', 'foot'),
      matchTexts
    },
    matchInputs,
    headerIndexes,
    larkContext: {
      accessToken,
      spreadsheetToken,
      sheetId,
      sheetName
    },
    sourceRowNumber: sourceRowEntry.rowNumber
  };
}

function coerceDateToYMD(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';

  const direct = parseDateYMD(raw);
  if (direct) return direct;

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return '';

  const y = parsed.getUTCFullYear();
  const m = String(parsed.getUTCMonth() + 1).padStart(2, '0');
  const d = String(parsed.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function resolveTeamIdFromTextSegment(text, teamsMap) {
  const normalizedText = normalizeOutcomeToken(text);
  if (!normalizedText) return '';

  let best = { teamId: '', score: 0 };
  for (const [teamId, team] of Object.entries(teamsMap)) {
    const candidates = [
      teamId,
      team.logo,
      team.en,
      team['zh-CN'],
      team['zh-TW'],
      team.ja
    ];

    for (const candidate of candidates) {
      if (!candidate) continue;
      const normalizedCandidate = normalizeOutcomeToken(candidate);
      if (!normalizedCandidate) continue;

      if (normalizedText.includes(normalizedCandidate) || normalizedCandidate.includes(normalizedText)) {
        if (normalizedCandidate.length > best.score) {
          best = { teamId, score: normalizedCandidate.length };
        }
      }

      const parts = String(candidate).trim().split(/\s+/);
      if (parts.length > 1) {
        const tail = normalizeOutcomeToken(parts[parts.length - 1]);
        if (tail && normalizedText.includes(tail) && tail.length > best.score) {
          best = { teamId, score: tail.length };
        }
      }
    }
  }

  return best.teamId;
}

function extractTeamIdsFromMatchQuestion(question, teamsMap) {
  const raw = String(question ?? '').trim();
  if (!raw) return [];

  const parts = raw.split(/\bvs\.?\b/i);
  if (parts.length < 2) return [];

  const leftRaw = parts[0]
    .replace(/^who\s+will\s+win[:\s-]*/i, '')
    .replace(/^will\s+/i, '')
    .trim();
  const rightRaw = parts.slice(1).join(' ')
    .replace(/\bon\b.+$/i, '')
    .replace(/\?+$/g, '')
    .trim();

  const leftId = resolveTeamIdFromTextSegment(leftRaw, teamsMap);
  const rightId = resolveTeamIdFromTextSegment(rightRaw, teamsMap);
  if (!leftId || !rightId || leftId === rightId) return [];

  return [leftId, rightId];
}

function pickNbaMoneylineMarketFromEvent(event) {
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  const filtered = markets.filter(market => {
    const slug = String(market?.slug ?? '').trim();
    if (!slug) return false;
    if (isPolymarketNbaMoneylineMarket(market)) return true;

    try {
      const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug);
      return Array.isArray(outcomes)
        && outcomes.length === 2
        && !isYesNoOutcomes(outcomes)
        && String(market?.question ?? '').toLowerCase().includes(' vs');
    } catch {
      return false;
    }
  });

  return filtered
    .sort((a, b) => Number(b?.volumeNum ?? b?.volume ?? 0) - Number(a?.volumeNum ?? a?.volume ?? 0))[0]
    ?? null;
}

function extractNbaDateFromMarket(market, event = null) {
  const candidates = [
    market?.gameStartTime,
    market?.eventStartTime,
    market?.startDateIso,
    market?.endDateIso,
    market?.startDate,
    market?.endDate,
    event?.gameStartTime,
    event?.eventStartTime,
    event?.startDateIso,
    event?.endDateIso,
    event?.startDate,
    event?.endDate,
    event?.ticker
  ];

  for (const candidate of candidates) {
    const ymd = coerceDateToYMD(candidate);
    if (ymd) return ymd;
  }
  return '';
}

// 足球专用日期提取：
// 1. 优先从 slug 末尾提取日期（如 epl-bur-mac-2026-04-22 → 2026-04-22），最准确
// 2. 次选 gameStartTime / eventStartTime / startDate
// 3. 最后才用 endDate（市场结算日，通常是比赛次日，会偏差 1 天）
function extractFootballMatchDate(market, event = null) {
  // 从 slug 末尾解析：xxx-YYYY-MM-DD
  const slug = String(market?.slug ?? event?.slug ?? '').trim();
  const slugDateMatch = slug.match(/(\d{4}-\d{2}-\d{2})$/);
  if (slugDateMatch) return slugDateMatch[1];

  const preferredCandidates = [
    market?.gameStartTime,
    market?.eventStartTime,
    market?.startDateIso,
    market?.startDate,
    event?.gameStartTime,
    event?.eventStartTime,
    event?.startDateIso,
    event?.startDate,
  ];
  for (const candidate of preferredCandidates) {
    const ymd = coerceDateToYMD(candidate);
    if (ymd) return ymd;
  }
  // endDate 作为最后兜底（可能偏差 1 天）
  const fallback = coerceDateToYMD(market?.endDateIso ?? market?.endDate ?? event?.endDateIso ?? event?.endDate ?? '');
  return fallback;
}

function resolveNbaTeamsFromMarket(market, teamsMap) {
  const slug = String(market?.slug ?? '').trim() || 'unknown';
  const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(item => String(item ?? '').trim());
  if (outcomes.length !== 2 || isYesNoOutcomes(outcomes)) {
    throw new Error(`NBA 市场解析失败（${slug}）：不是双边球队盘口`);
  }

  const questionTeams = extractTeamIdsFromMatchQuestion(market?.question ?? '', teamsMap);
  if (questionTeams.length === 2) {
    return {
      homeTeamId: questionTeams[0],
      awayTeamId: questionTeams[1]
    };
  }

  const outcomeTeamIds = outcomes.map(outcome => resolveTeamId(outcome, teamsMap));
  if (outcomeTeamIds[0] && outcomeTeamIds[1] && outcomeTeamIds[0] !== outcomeTeamIds[1]) {
    return {
      homeTeamId: outcomeTeamIds[0],
      awayTeamId: outcomeTeamIds[1]
    };
  }

  throw new Error(`NBA 球队解析失败（${slug}）：question=${market?.question ?? ''}, outcomes=${JSON.stringify(outcomes)}`);
}

async function resolveClassicMatchFromPolymarketInput(matchInput, teamsMap) {
  const rawUrl = String(matchInput?.polymarket_url ?? '').trim();
  const slug = normalizePolymarketInputSlug(rawUrl);
  if (!slug) {
    throw new Error(`第 ${matchInput?.index ?? '?'} 场比赛的 Polymarket 链接无效：${rawUrl}`);
  }

  let market = null;
  let event = null;
  try {
    market = await fetchPolymarketMarketBySlug(slug);
  } catch (err) {
    if (!String(err?.message ?? '').includes('HTTP 404')) throw err;
  }

  if (!market) {
    event = await fetchPolymarketEventBySlug(slug);
    market = pickNbaMoneylineMarketFromEvent(event);
    if (!market) {
      throw new Error(`第 ${matchInput?.index ?? '?'} 场比赛未找到可用 NBA moneyline 市场：${slug}`);
    }
  }

  const { homeTeamId, awayTeamId } = resolveNbaTeamsFromMarket(market, teamsMap);
  const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', String(market?.slug ?? slug)).map(String);
  const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', String(market?.slug ?? slug));
  if (outcomes.length !== prices.length) {
    throw new Error(`第 ${matchInput?.index ?? '?'} 场比赛的 Polymarket 数据异常：outcomes 与 outcomePrices 数量不一致`);
  }

  const homeIndex = findOutcomeIndexByAliases(outcomes, buildTeamAliasTokens(homeTeamId, teamsMap));
  const awayIndex = findOutcomeIndexByAliases(outcomes, buildTeamAliasTokens(awayTeamId, teamsMap));
  if (homeIndex === -1 || awayIndex === -1 || homeIndex === awayIndex) {
    throw new Error(`第 ${matchInput?.index ?? '?'} 场比赛的 outcome 映射失败：${slug}`);
  }

  const rounded = roundWinRatesToIntegers(
    toPercentProbability(prices[homeIndex]),
    toPercentProbability(prices[awayIndex])
  );
  if (!Number.isFinite(rounded.homeWin) || !Number.isFinite(rounded.awayWin)) {
    throw new Error(`第 ${matchInput?.index ?? '?'} 场比赛的赔率解析失败：${slug}`);
  }

  const date = extractNbaDateFromMarket(market, event);
  return {
    index: Number(matchInput?.index ?? 0),
    date,
    home_team: homeTeamId,
    away_team: awayTeamId,
    home_win: String(rounded.homeWin),
    away_win: String(rounded.awayWin),
    polymarket_url: rawUrl || buildPolymarketEventUrl(String(event?.slug ?? market?.slug ?? slug)),
    polymarket_slug: String(market?.slug ?? slug).trim()
  };
}

async function writeBackClassicMatchesToLark(gamesRows, headerIndexes, larkContext, sourceRowNumber) {
  const writableFieldMap = {
    1: ['match1_home', 'match1_away', 'match1_date', 'match1_home_win', 'match1_away_win'],
    2: ['match2_home', 'match2_away', 'match2_date', 'match2_home_win', 'match2_away_win'],
    3: ['match3_home', 'match3_away', 'match3_date', 'match3_home_win', 'match3_away_win']
  };

  const availableFields = new Set(
    Object.values(writableFieldMap)
      .flat()
      .filter(field => Number.isInteger(headerIndexes[field]))
  );
  if (availableFields.size === 0) {
    return 0;
  }

  for (const row of gamesRows) {
    const index = Number(row.index);
    const updates = [
      [`match${index}_home`, row.home_team],
      [`match${index}_away`, row.away_team],
      [`match${index}_date`, row.date],
      [`match${index}_home_win`, formatPercentForSheet(row.home_win)],
      [`match${index}_away_win`, formatPercentForSheet(row.away_win)]
    ];

    for (const [field, value] of updates) {
      if (!availableFields.has(field)) continue;
      const colA1 = indexToColumnLabel(headerIndexes[field]);
      await updateLarkCellValue({
        accessToken: larkContext.accessToken,
        spreadsheetToken: larkContext.spreadsheetToken,
        sheetId: larkContext.sheetId,
        a1Cell: `${colA1}${sourceRowNumber}`,
        value
      });
    }
  }

  return gamesRows.length;
}

async function translateClassicTitles(sourceData, targetLangs, fromLang = 'zh-CN') {
  const matchTexts = Array.isArray(sourceData.matchTexts) ? sourceData.matchTexts : [];
  const baseTexts = [sourceData.mainTitle, sourceData.subTitle, sourceData.footer];
  const matchFlat = [];
  for (const mt of matchTexts) {
    matchFlat.push(mt?.reward ?? '', mt?.news ?? '');
  }
  const allTexts = [...baseTexts, ...matchFlat];
  const result = {};

  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(allTexts.map(text => translateOneText(text, fromLang, lang)));
    const translatedMatchTexts = [];
    for (let i = 0; i < matchTexts.length; i++) {
      translatedMatchTexts.push({
        reward: translated[3 + i * 2] ?? '',
        news: translated[3 + i * 2 + 1] ?? ''
      });
    }
    result[lang] = {
      mainTitle: translated[0],
      subTitle: translated[1],
      footer: translated[2],
      matchTexts: translatedMatchTexts
    };
    console.log(' ✅');
  }

  return result;
}

async function writeBackClassicTranslationsToLark(sourceData, translationsMap, accessToken, spreadsheetToken, sheetId, sourceLang = 'zh-CN', headerIndexes = {}) {
  const sourceLangNormalized = String(sourceLang ?? '').trim().toLowerCase();
  const langs = Object.keys(translationsMap)
    .filter(lang => String(lang ?? '').trim().toLowerCase() !== sourceLangNormalized);

  if (langs.length === 0) return 0;

  const translatedRows = langs.map(lang => {
    const d = translationsMap[lang] ?? {};
    return [
      lang,
      String(d.mainTitle ?? ''),
      String(d.subTitle ?? ''),
      String(d.footer ?? '')
    ];
  });

  const CLEAR_WINDOW_ROWS = 17; // rows 3..19
  const blankRow = ['', '', '', ''];
  const rows = Array.from({ length: CLEAR_WINDOW_ROWS }, (_, i) => translatedRows[i] ?? blankRow);
  const startRow = 3;
  const endRow = startRow + rows.length - 1;
  const range = `${sheetId}!A${startRow}:D${endRow}`;

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({ valueRange: { range, values: rows } })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`NBA 翻译回填 Lark 失败：${data.msg || res.status}`);
  }

  // 回填每场比赛的 reward/news 翻译（位置与源行同列，按语种逐行写入）
  const rewardNewsFields = [];
  for (let i = 1; i <= 3; i++) {
    rewardNewsFields.push(
      { key: `match${i}_reward`, matchIdx: i - 1, type: 'reward' },
      { key: `match${i}_news`, matchIdx: i - 1, type: 'news' }
    );
  }

  for (let langIdx = 0; langIdx < langs.length; langIdx++) {
    const lang = langs[langIdx];
    const rowNumber = startRow + langIdx;
    const translated = translationsMap[lang] ?? {};
    const matchTexts = Array.isArray(translated.matchTexts) ? translated.matchTexts : [];

    for (const field of rewardNewsFields) {
      const colIdx = headerIndexes[field.key];
      if (!Number.isInteger(colIdx)) continue;
      const value = String(matchTexts[field.matchIdx]?.[field.type] ?? '');
      const colA1 = indexToColumnLabel(colIdx);
      await updateLarkCellValue({
        accessToken,
        spreadsheetToken,
        sheetId,
        a1Cell: `${colA1}${rowNumber}`,
        value
      });
    }
  }

  return translatedRows.length;
}

function normalizeOutcomeToken(value) {
  return String(value ?? '')
    .normalize('NFKD')
    .replace(/\p{M}+/gu, '')
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

function parseJsonArrayField(value, fieldName, slug) {
  if (Array.isArray(value)) return value;
  if (typeof value !== 'string') {
    throw new Error(`Polymarket 字段异常（${slug}.${fieldName}）：不是数组或 JSON 字符串`);
  }
  try {
    const parsed = JSON.parse(value);
    if (!Array.isArray(parsed)) {
      throw new Error('parsed value is not array');
    }
    return parsed;
  } catch {
    throw new Error(`Polymarket 字段异常（${slug}.${fieldName}）：JSON 解析失败`);
  }
}

function sanitizeCardText(text) {
  const raw = String(text ?? '').trim();
  if (!raw) return '';
  return raw.replace(/^#+\s*/u, '').trim();
}

function buildTeamAliasTokens(teamId, teamsMap) {
  const team = teamsMap[teamId] ?? {};
  const candidates = [
    teamId,
    team.logo,
    team.en,
    team['zh-CN'],
    team['zh-TW'],
    team.ja
  ];

  const aliases = new Set();
  for (const candidate of candidates) {
    if (!candidate) continue;
    const normalized = normalizeOutcomeToken(candidate);
    if (normalized) aliases.add(normalized);

    // 英文名可能带城市名，补一个最后单词作为兜底（如 "Los Angeles Lakers" -> "Lakers"）
    const parts = String(candidate).trim().split(/\s+/);
    if (parts.length > 1) {
      const tail = normalizeOutcomeToken(parts[parts.length - 1]);
      if (tail) aliases.add(tail);
    }
  }
  return aliases;
}

function resolveTeamId(inputTeamId, teamsMap) {
  const raw = String(inputTeamId ?? '').trim();
  if (!raw) return raw;

  if (teamsMap[raw]) return raw;

  const lowered = raw.toLowerCase();
  if (teamsMap[lowered]) return lowered;

  const normalizedRaw = normalizeOutcomeToken(raw);
  if (!normalizedRaw) return raw;

  // 先做精确匹配 + 候选尾词匹配
  for (const [teamId, team] of Object.entries(teamsMap)) {
    const candidates = [
      teamId,
      team.logo,
      team.en,
      team['zh-CN'],
      team['zh-TW'],
      team.ja
    ];
    for (const candidate of candidates) {
      if (!candidate) continue;
      const normalizedCandidate = normalizeOutcomeToken(candidate);
      if (normalizedCandidate && normalizedCandidate === normalizedRaw) {
        return teamId;
      }

      const parts = String(candidate).trim().split(/\s+/);
      if (parts.length > 1) {
        const tail = normalizeOutcomeToken(parts[parts.length - 1]);
        if (tail && tail === normalizedRaw) {
          return teamId;
        }
      }
    }
  }

  // 输入本身是多词（如 "Atlanta Hawks"）时，用输入的尾词反向匹配
  // 处理 Polymarket outcomes 里写完整城市+队名的情况
  const inputParts = raw.trim().split(/\s+/);
  if (inputParts.length > 1) {
    const inputTail = normalizeOutcomeToken(inputParts[inputParts.length - 1]);
    if (inputTail) {
      // 精确匹配 id
      if (teamsMap[inputTail]) return inputTail;
      // 匹配 en 字段
      for (const [teamId, team] of Object.entries(teamsMap)) {
        if (normalizeOutcomeToken(team.en) === inputTail) return teamId;
      }
    }
  }

  // 英超及其他手动别名兜底
  const TEAM_ALIASES = [
    ['tottenham', ['spurs', 'tottenhamhotspur', 'tottenhamhotspurfc']],
    ['wolves', ['wolverhampton', 'wolverhamptonwanderers', 'wolverhamptonwanderersfc']],
    ['west_ham', ['westham', 'westhamunited', 'westhamunitedfc']],
    ['nottingham_forest', ['forest', 'nottinghamforest', 'nottinghamforestfc']],
    ['brighton', ['brightonandhovealbion', 'brightonandhove', 'brightonhovealbion']],
    ['newcastle', ['newcastleunited', 'newcastleunitedfc', 'newcastleutd']],
    ['bournemouth', ['afcbournemouth']],
    ['leeds_united', ['leeds', 'leedsunited', 'leedsunitedfc']],
    ['sunderland', ['sunderlandafc']],
    ['crystal_palace', ['palace', 'crystalpalacefc']],
    ['aston_villa', ['astonvillafc', 'villa']],
    ['brentford', ['brentfordfc']],
    ['chelsea', ['chelseafc', 'cfc']],
    ['everton', ['evertonfc', 'toffees']],
    ['fulham', ['fulhamfc']],
    ['burnley', ['burnleyfc']],
    ['liverpool', ['liverpoolfc', 'lfc']],
    ['arsenal', ['arsenalfc', 'gunners']],
    ['man_city', ['manchestercity', 'mancity', 'manchestercityfc', 'mcfc', 'cityfc', 'city']],
    ['man_utd', ['manchesterunited', 'manutd', 'manchesterunitedfc', 'mufc', 'unitedfc', 'manunited']],
    // 其他联赛常见别名
    ['psg', ['parissaintgermainfc', 'psg', 'parissaintgermain']],
    ['atletico_madrid', ['clubatleticodemadrid', 'atleticomadrid', 'atletico']],
    ['athletic_bilbao', ['athleticclub', 'athleticclubbilbao']],
    ['inter_milan', ['inter', 'internazionale', 'fcinternazionale']],
    ['ac_milan', ['milan']],
    ['sporting_cp', ['sporting', 'sportinglisbon']],
    ['usa', ['unitedstates', 'us']],
    ['south_korea', ['korearepublic', 'korea']],
    ['netherlands', ['holland']],
  ];
  for (const [teamId, aliases] of TEAM_ALIASES) {
    if (aliases.includes(normalizedRaw) && teamsMap[teamId]) return teamId;
  }

  return raw;
}

function normalizeGameRowsTeamIds(gamesRows, teamsMap) {
  return gamesRows.map((row, index) => {
    const normalizedHome = resolveTeamId(row.home_team, teamsMap);
    const normalizedAway = resolveTeamId(row.away_team, teamsMap);

    if (normalizedHome !== row.home_team) {
      warnOnce(
        `team-normalized-home:${index}:${row.home_team}`,
        `已自动归一主队 ID：${row.home_team} -> ${normalizedHome}`
      );
    }
    if (normalizedAway !== row.away_team) {
      warnOnce(
        `team-normalized-away:${index}:${row.away_team}`,
        `已自动归一客队 ID：${row.away_team} -> ${normalizedAway}`
      );
    }

    return {
      ...row,
      home_team: normalizedHome,
      away_team: normalizedAway
    };
  });
}

function findOutcomeIndexByAliases(outcomes, aliases) {
  if (!aliases || aliases.size === 0) return -1;
  const normalizedOutcomes = outcomes.map(item => normalizeOutcomeToken(item));

  const exactIdx = normalizedOutcomes.findIndex(item => aliases.has(item));
  if (exactIdx !== -1) return exactIdx;

  for (let i = 0; i < normalizedOutcomes.length; i++) {
    const outcome = normalizedOutcomes[i];
    for (const alias of aliases) {
      if (!alias || alias.length < 3 || outcome.length < 3) continue;
      if (outcome.includes(alias) || alias.includes(outcome)) {
        return i;
      }
    }
  }

  return -1;
}

function toPercentProbability(raw) {
  const value = Number(raw);
  if (!Number.isFinite(value)) return NaN;
  if (value >= 0 && value <= 1) return value * 100;
  return value;
}

function round2(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function roundWinRatesToIntegers(homeRaw, awayRaw) {
  const home = Number(homeRaw);
  const away = Number(awayRaw);
  if (!Number.isFinite(home) || !Number.isFinite(away)) {
    return { homeWin: NaN, awayWin: NaN };
  }

  let homeWin = Math.round(home);
  let awayWin = Math.round(away);
  const sum = homeWin + awayWin;

  if (sum !== 100) {
    const homeDelta = Math.abs(home - homeWin);
    const awayDelta = Math.abs(away - awayWin);

    if (sum > 100) {
      if (homeDelta >= awayDelta) {
        homeWin -= sum - 100;
      } else {
        awayWin -= sum - 100;
      }
    } else {
      if (homeDelta >= awayDelta) {
        homeWin += 100 - sum;
      } else {
        awayWin += 100 - sum;
      }
    }
  }

  return { homeWin, awayWin };
}

function extractPolymarketSlug(input) {
  const raw = String(input ?? '').trim();
  if (!raw) return '';

  try {
    const url = new URL(raw);
    const parts = url.pathname.split('/').filter(Boolean);
    const eventIndex = parts.findIndex(part => part === 'event' || part === 'market');
    if (eventIndex !== -1 && parts[eventIndex + 1]) {
      return String(parts[eventIndex + 1]).trim();
    }
    const candidate = parts[parts.length - 1];
    if (candidate) return String(candidate).trim();
  } catch {
    return raw.replace(/^\/+|\/+$/g, '');
  }

  return raw.replace(/^\/+|\/+$/g, '');
}

function normalizePolymarketInputSlug(rawInput) {
  const raw = String(rawInput ?? '').trim();
  if (!raw) return '';

  const normalized = extractPolymarketSlug(raw);
  // 周赛列表链接会解析成纯数字（如 "26"），不可直接用于抓赔率，回退自动匹配。
  if (/^\d+$/.test(normalized) && /\/sports\/.+\/games\/week\//i.test(raw)) {
    return '';
  }
  return normalized;
}

function getRowPolymarketInput(row) {
  return String(row?.polymarket_url ?? row?.polymarket_slug ?? '').trim();
}

function buildPolymarketEventUrl(slug) {
  const normalized = String(slug ?? '').trim();
  return normalized ? `https://polymarket.com/event/${normalized}` : '';
}

function normalizeConfigKey(key) {
  return String(key ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_');
}

function parseComprehensiveScenarioConfig(values, startRow = 20, endRow = 35) {
  const config = {};
  for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
    const row = values[rowNum - 1] ?? [];
    const key = normalizeConfigKey(row[0]);
    const value = String(row[1] ?? '').trim();
    if (!key) continue;
    config[key] = value;
  }
  return config;
}

function resolveWorldCupTeamByOutcome(outcomeText, footballTeamsMap) {
  const normalizedOutcome = normalizeOutcomeToken(outcomeText);
  if (!normalizedOutcome) return null;

  for (const team of Object.values(footballTeamsMap)) {
    const aliases = [
      team?.id,
      team?.en,
      team?.['zh-CN'],
      team?.['zh-TW'],
      team?.ja
    ]
      .map(item => normalizeOutcomeToken(item))
      .filter(Boolean);

    if (aliases.some(alias => alias === normalizedOutcome)) {
      return team;
    }
  }

  const manualAliases = [
    // 国家队
    ['usa', ['unitedstates', 'unitedstatesofamerica', 'us']],
    ['south_korea', ['korearepublic', 'republicofkorea', 'korea']],
    ['england', ['englandnationalteam']],
    ['netherlands', ['holland']],
    ['ivory_coast', ['cotedivoire']],
    // 法甲
    ['psg', ['parissaintgermainfc', 'psg', 'parissaintgermain']],
    ['lyon', ['olympiquelyonnais', 'olympiquelyon', 'ol']],
    ['marseille', ['olympiquemarseille', 'om']],
    ['rennes', ['staderennais', 'rennais']],
    ['brest', ['stadebrestois', 'brestois']],
    // 西甲
    ['atletico_madrid', ['clubatleticodemadrid', 'atleticomadrid', 'atleticodemadrid', 'atleticosdemadrid', 'atletico']],
    ['celta_vigo', ['celtadevigo', 'rccelta']],
    ['athletic_bilbao', ['athleticclub', 'athleticclubbilbao', 'athleticclubdebilbao']],
    // 德甲
    ['bayern', ['bayernmunchen', 'bayernmunich', 'fcbayern', 'fcbayernmunich']],
    ['monchengladbach', ['borussiamonchengladbach', 'monchengladbach', 'mgladbach']],
    ['hamburg', ['hamburger', 'hamburgsv']],
    ['cologne', ['fckoln', 'koln', 'fckolnde', '1fckoln']],
    // 意甲
    ['inter_milan', ['inter', 'internazionale', 'fcinternazionale', 'fcinter']],
    ['ac_milan', ['milan']],
    ['juventus', ['juve']],
    // 葡超
    ['sporting_cp', ['sporting', 'sportinglisbon', 'sportingportugal']],
    // 英超
    ['tottenham', ['spurs', 'tottenhamhotspur', 'tottenhamhotspurfc']],
    ['wolves', ['wolverhampton', 'wolverhamptonwanderers', 'wolverhamptonwanderersfc']],
    ['west_ham', ['westham', 'westhamunited', 'westhamunitedfc']],
    ['nottingham_forest', ['forest', 'nottinghamforest', 'nottinghamforestfc']],
    ['brighton', ['brightonandhovealbion', 'brightonandhove', 'brightonhovealbion']],
    ['newcastle', ['newcastleunited', 'newcastleunitedfc', 'newcastleutd']],
    ['bournemouth', ['afcbournemouth']],
    ['leeds_united', ['leeds', 'leedsunited', 'leedsunitedfc']],
    ['sunderland', ['sunderlandafc']],
    ['crystal_palace', ['palace', 'crystalpalacefc']],
    ['aston_villa', ['astonvillafc', 'villa']],
    ['brentford', ['brentfordfc']],
    ['chelsea', ['chelseafc', 'cfc']],
    ['everton', ['evertonfc', 'toffees']],
    ['fulham', ['fulhamfc']],
    ['burnley', ['burnleyfc']],
    ['liverpool', ['liverpoolfc', 'lfc']],
    ['arsenal', ['arsenalfc', 'gunners']],
    ['man_city', ['manchestercity', 'mancity', 'manchestercityfc', 'mcfc', 'cityfc', 'city']],
    ['man_utd', ['manchesterunited', 'manutd', 'manchesterunitedfc', 'mufc', 'unitedfc', 'manunited']],
  ];

  for (const [teamId, aliases] of manualAliases) {
    if (!aliases.includes(normalizedOutcome)) continue;
    if (footballTeamsMap[teamId]) return footballTeamsMap[teamId];
  }

  return null;
}

function buildWorldCupCardsFromPolymarketMarket(market, cardCount, footballTeamsMap) {
  const slug = String(market?.slug ?? '').trim() || 'unknown';
  const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(item => String(item ?? '').trim());
  const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', slug);
  if (!Array.isArray(outcomes) || !Array.isArray(prices) || outcomes.length !== prices.length) {
    throw new Error(`Polymarket 市场数据异常（${slug}）：outcomes 与 outcomePrices 不匹配`);
  }

  const cards = outcomes
    .map((outcome, index) => {
      const percent = toPercentProbability(prices[index]);
      if (!Number.isFinite(percent)) return null;
      const rounded = Math.max(0, Math.min(100, Math.round(percent)));
      const team = resolveWorldCupTeamByOutcome(outcome, footballTeamsMap);
      const displayText = sanitizeCardText(team?.['zh-CN'] ?? outcome);

      let image = '';
      if (team?.logo) {
        const rawLogoPath = String(team.logo).trim();
        const absLogoPath = path.isAbsolute(rawLogoPath) ? rawLogoPath : path.join(BASE_DIR, rawLogoPath);
        if (fs.existsSync(absLogoPath)) {
          image = `file://${absLogoPath}`;
        }
      }

      return {
        text: displayText,
        percent: rounded,
        image
      };
    })
    .filter(Boolean)
    .sort((a, b) => b.percent - a.percent)
    .slice(0, Math.max(2, Math.min(4, Number(cardCount) || 3)));

  return cards;
}

function extractTeamNameFromQuestion(question = '') {
  const raw = String(question ?? '').trim();
  const match = raw.match(/^Will\s+(.+?)\s+win\b/i);
  if (match && match[1]) return String(match[1]).trim();
  return raw;
}

function buildWorldCupCardsFromPolymarketEvent(event, cardCount, footballTeamsMap) {
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  const yesAliases = new Set(['yes', 'y']);
  const cards = [];

  for (const market of markets) {
    const slug = String(market?.slug ?? '').trim() || String(event?.slug ?? 'event');
    const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(item => String(item ?? '').trim());
    const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', slug);
    if (!Array.isArray(outcomes) || !Array.isArray(prices) || outcomes.length !== prices.length) continue;

    let yesIndex = outcomes.findIndex(outcome => yesAliases.has(normalizeOutcomeToken(outcome)));
    if (yesIndex === -1 && outcomes.length === 2) {
      yesIndex = 0;
    }
    if (yesIndex === -1) continue;

    const percent = toPercentProbability(prices[yesIndex]);
    if (!Number.isFinite(percent)) continue;
    const rounded = Math.max(0, Math.min(100, Math.round(percent)));

    const teamLabel = String(market?.groupItemTitle ?? '').trim() || extractTeamNameFromQuestion(market?.question ?? '');
    if (!teamLabel) continue;

    const team = resolveWorldCupTeamByOutcome(teamLabel, footballTeamsMap);
    const displayText = sanitizeCardText(team?.['zh-CN'] ?? teamLabel);

    let image = '';
    if (team?.logo) {
      const rawLogoPath = String(team.logo).trim();
      const absLogoPath = path.isAbsolute(rawLogoPath) ? rawLogoPath : path.join(BASE_DIR, rawLogoPath);
      if (fs.existsSync(absLogoPath)) {
        image = `file://${absLogoPath}`;
      }
    }

    cards.push({ text: displayText, percent: rounded, image });
  }

  return cards
    .sort((a, b) => b.percent - a.percent)
    .slice(0, Math.max(2, Math.min(4, Number(cardCount) || 3)));
}

async function buildWorldCupCardsFromPolymarketInput(inputSlugOrUrl, cardCount, footballTeamsMap) {
  const slug = extractPolymarketSlug(inputSlugOrUrl);
  if (!slug) return [];

  try {
    const market = await fetchPolymarketMarketBySlug(slug);
    return buildWorldCupCardsFromPolymarketMarket(market, cardCount, footballTeamsMap);
  } catch (err) {
    const message = String(err?.message ?? '');
    if (!message.includes('HTTP 404')) throw err;
  }

  const event = await fetchPolymarketEventBySlug(slug);
  return buildWorldCupCardsFromPolymarketEvent(event, cardCount, footballTeamsMap);
}

// Extract home/away team names and odds from a Polymarket football market or event
async function buildFootballGameRowFromPolymarketLink(slugOrUrl, teamsMap) {
  const slug = normalizePolymarketInputSlug(slugOrUrl);
  if (!slug) throw new Error(`无效的 Polymarket 链接：${slugOrUrl}`);

  let market = null;
  let event = null;
  try {
    market = await fetchPolymarketMarketBySlug(slug);
  } catch (err) {
    if (!String(err?.message ?? '').includes('HTTP 404')) throw err;
    event = await fetchPolymarketEventBySlug(slug);
  }

  if (market) {
    const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(s => String(s ?? '').trim());
    const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', slug);

    // 检测平局 outcome：匹配 "draw"、"tie"，也兼容 Polymarket 的 "Draw (Team A vs Team B)" 格式
    const isDrawOutcome = (o) => {
      const n = normalizeOutcomeToken(o);
      return n === 'draw' || n === 'tie' || n.startsWith('draw') || n.startsWith('tie');
    };

    // Determine home/away from question ("Team A vs Team B")
    const question = String(market?.question ?? '').trim();
    const vsMatch = question.match(/^(.+?)\s+vs\.?\s+(.+?)(?:\?|$)/i);
    let homeRaw, awayRaw;
    if (vsMatch) {
      homeRaw = vsMatch[1].trim();
      awayRaw = vsMatch[2].trim();
    } else {
      // 从 Draw outcome 标签提取正确主客队顺序："Draw (Home vs. Away)"
      // Polymarket 足球 3 结果市场的 Draw label 始终保持 "主队 vs 客队" 顺序
      const drawOutcome = outcomes.find(o => isDrawOutcome(o)) ?? '';
      const drawVsMatch = String(drawOutcome).match(/\((.+?)\s+vs\.?\s+(.+?)\)/i);
      if (drawVsMatch) {
        homeRaw = drawVsMatch[1].trim();
        awayRaw = drawVsMatch[2].trim();
      } else {
        // 最后兜底：按 outcomes 列表顺序（可能不准确）
        homeRaw = outcomes.find(o => !isDrawOutcome(o)) ?? outcomes[0] ?? '';
        awayRaw = outcomes.filter(o => !isDrawOutcome(o))[1] ?? outcomes[1] ?? '';
      }
    }

    const toOdds = (name) => {
      const idx = outcomes.findIndex(o => normalizeOutcomeToken(o) === normalizeOutcomeToken(name));
      return idx >= 0 ? Math.round(toPercentProbability(prices[idx])) : 0;
    };
    const drawIdx = outcomes.findIndex(o => isDrawOutcome(o));
    const drawWin = drawIdx >= 0 ? Math.round(toPercentProbability(prices[drawIdx])) : null;

    const resolvedHome = resolveWorldCupTeamByOutcome(homeRaw, teamsMap);
    const resolvedAway = resolveWorldCupTeamByOutcome(awayRaw, teamsMap);

    // 日期：优先从 slug 末尾提取（最准确），再兜底用 API 字段
    const slugDateMatch = slug.match(/(\d{4}-\d{2}-\d{2})$/);
    const matchDate = slugDateMatch ? slugDateMatch[1] : extractFootballMatchDate(market, event);

    return {
      date: matchDate,
      home_team: resolvedHome?.id ?? homeRaw,
      away_team: resolvedAway?.id ?? awayRaw,
      home_win: String(toOdds(homeRaw)),
      away_win: String(toOdds(awayRaw)),
      draw_win: drawWin !== null ? String(drawWin) : '',
      polymarket_slug: slug
    };
  }

  // Event (multiple yes/no markets per team)
  // Polymarket 足球事件结构：每个 sub-market 对应一个结果（主队赢/平局/客队赢），各自是 Yes/No 市场
  const markets = Array.isArray(event?.markets) ? event.markets : [];
  const yesAliases = new Set(['yes', 'y']);

  // 检测平局 label（同上 isDrawOutcome 逻辑）
  const isDrawLabel = (label) => {
    const n = normalizeOutcomeToken(label);
    return n === 'draw' || n === 'tie' || n.startsWith('draw') || n.startsWith('tie');
  };

  const teamEntries = [];
  let evDrawPercent = null;
  let evDrawLabel = '';

  for (const m of markets) {
    const mOutcomes = parseJsonArrayField(m?.outcomes ?? '[]', 'outcomes', slug).map(s => String(s ?? '').trim());
    const mPrices = parseJsonArrayField(m?.outcomePrices ?? '[]', 'outcomePrices', slug);
    const yesIdx = mOutcomes.findIndex(o => yesAliases.has(normalizeOutcomeToken(o)));
    if (yesIdx === -1) continue;
    const percent = Math.round(toPercentProbability(mPrices[yesIdx]));
    const label = String(m?.groupItemTitle ?? '').trim() || extractTeamNameFromQuestion(m?.question ?? '');
    if (!label) continue;

    if (isDrawLabel(label)) {
      // 这是平局子市场：存平局概率，从 Draw label 中提取主客队顺序
      evDrawPercent = percent;
      evDrawLabel = label;
    } else {
      teamEntries.push({ label, percent });
    }
  }

  // 尝试从 Draw label 中提取主客队顺序（最可靠）
  let evHomeRaw = '', evAwayRaw = '';
  if (evDrawLabel) {
    const drawVsMatch = evDrawLabel.match(/\((.+?)\s+vs\.?\s+(.+?)\)/i);
    if (drawVsMatch) {
      evHomeRaw = drawVsMatch[1].trim();
      evAwayRaw = drawVsMatch[2].trim();
    }
  }

  if (!evHomeRaw) {
    // 兜底：按胜率降序（但主客队顺序可能不准）
    teamEntries.sort((a, b) => b.percent - a.percent);
    evHomeRaw = teamEntries[0]?.label ?? '';
    evAwayRaw = teamEntries[1]?.label ?? '';
  } else {
    // 用从 Draw label 提取的主客队名查找对应胜率
    const findPct = (rawName) => {
      const n = normalizeOutcomeToken(rawName);
      const entry = teamEntries.find(e => normalizeOutcomeToken(e.label) === n);
      if (entry) return entry.percent;
      // 宽松匹配：any entry whose normalized label contains the rawName's normalized tail
      const tail = rawName.trim().split(/\s+/).pop() ?? '';
      const tailN = normalizeOutcomeToken(tail);
      return teamEntries.find(e => normalizeOutcomeToken(e.label).includes(tailN))?.percent ?? 0;
    };
    // re-sort by home/away label
    const homePct = findPct(evHomeRaw);
    const awayPct = findPct(evAwayRaw);
    // assign correctly
    evHomeRaw = evHomeRaw; // already set
    evAwayRaw = evAwayRaw;
    teamEntries[0] = { label: evHomeRaw, percent: homePct };
    teamEntries[1] = { label: evAwayRaw, percent: awayPct };
  }

  const homeEntry = teamEntries[0];
  const awayEntry = teamEntries[1];
  if (!homeEntry || !awayEntry) throw new Error(`无法从 Polymarket 事件解析球队数据（${slug}）`);

  const resolvedHome = resolveWorldCupTeamByOutcome(homeEntry.label, teamsMap);
  const resolvedAway = resolveWorldCupTeamByOutcome(awayEntry.label, teamsMap);

  const slugDateMatchEv = slug.match(/(\d{4}-\d{2}-\d{2})$/);
  const matchDateEv = slugDateMatchEv ? slugDateMatchEv[1] : extractFootballMatchDate(null, event);

  return {
    date: matchDateEv,
    home_team: resolvedHome?.id ?? homeEntry.label,
    away_team: resolvedAway?.id ?? awayEntry.label,
    home_win: String(homeEntry.percent),
    away_win: String(awayEntry.percent),
    draw_win: evDrawPercent !== null ? String(evDrawPercent) : '',
    polymarket_slug: slug
  };
}

// Resolve link_only rows by fetching data from Polymarket
async function resolveFootballLinkOnlyRows(rows, teamsMap) {
  const resolved = [];
  for (const row of rows) {
    // 若 away_team 或 home_team 包含 Draw 开头的字符串（上次解析写坏），强制重新解析
    const hasCorruptedTeam = (t) => /^draw[\s(]/i.test(String(t ?? '').trim());
    const needsReFetch = row.link_only ||
      (row.polymarket_slug && (hasCorruptedTeam(row.home_team) || hasCorruptedTeam(row.away_team)));
    if (!needsReFetch) {
      resolved.push(row);
      continue;
    }
    try {
      process.stdout.write(`  → 从链接获取比赛数据（${row.polymarket_slug}）...`);
      const gameRow = await buildFootballGameRowFromPolymarketLink(row.polymarket_slug, teamsMap);
      // preserve manually set date if present
      if (row.date) gameRow.date = row.date;
      console.log(` ✅ ${gameRow.home_team} vs ${gameRow.away_team}`);
      resolved.push(gameRow);
    } catch (err) {
      console.log(` ❌ ${err.message}`);
    }
  }
  return resolved;
}

async function writeBackComprehensiveCardsToLark(cards, accessToken, spreadsheetToken, sheetId) {
  const MAX_CARDS = 4;
  const row = [];
  for (let i = 0; i < MAX_CARDS; i++) {
    const card = cards[i];
    row.push(String(card?.text ?? ''), card?.percent ?? '');
  }

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({
      valueRange: {
        range: `${sheetId}!E2:L2`,
        values: [row]
      }
    })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`回填世界杯卡片失败：${data.msg || res.status}`);
  }
}

async function fetchPolymarketMarketBySlug(slug) {
  const url = `${POLYMARKET_BASE_URL}/markets/slug/${encodeURIComponent(slug)}`;
  const res = await fetch(url);
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    const detail = text ? `，响应：${text.slice(0, 120)}` : '';
    throw new Error(`Polymarket 读取失败（${slug}）：HTTP ${res.status}${detail}`);
  }
  return res.json();
}

async function fetchPolymarketEventBySlug(slug) {
  const url = `${POLYMARKET_BASE_URL}/events/slug/${encodeURIComponent(slug)}`;
  const res = await fetch(url);
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    const detail = text ? `，响应：${text.slice(0, 120)}` : '';
    throw new Error(`Polymarket 事件读取失败（${slug}）：HTTP ${res.status}${detail}`);
  }
  return res.json();
}

async function fetchPolymarketJson(url, label) {
  const res = await fetch(url);
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    const detail = text ? `，响应：${text.slice(0, 120)}` : '';
    throw new Error(`Polymarket ${label}失败：HTTP ${res.status}${detail}`);
  }
  return res.json();
}

async function fetchPolymarketSportsMetadata() {
  return fetchPolymarketJson(`${POLYMARKET_BASE_URL}/sports`, 'sports metadata 读取');
}

function parseNumericIdList(value) {
  return String(value ?? '')
    .split(',')
    .map(item => Number(String(item).trim()))
    .filter(item => Number.isFinite(item) && item > 0);
}

function resolvePolymarketNbaTagId(sportsMetadata) {
  const sports = Array.isArray(sportsMetadata) ? sportsMetadata : [];
  const nbaMeta = sports.find(item => normalizeOutcomeToken(item?.sport) === 'nba');
  if (!nbaMeta) return 745;

  const frequency = new Map();
  for (const sport of sports) {
    for (const tagId of parseNumericIdList(sport?.tags)) {
      frequency.set(tagId, (frequency.get(tagId) ?? 0) + 1);
    }
  }

  const candidateTags = parseNumericIdList(nbaMeta.tags).filter(tagId => tagId !== 1);
  if (candidateTags.length === 0) return 745;

  const ranked = candidateTags
    .map(tagId => ({ tagId, frequency: frequency.get(tagId) ?? Number.MAX_SAFE_INTEGER }))
    .sort((a, b) => a.frequency - b.frequency || a.tagId - b.tagId);

  return ranked[0]?.tagId ?? 745;
}

function buildPolymarketMarketsUrl(tagId, offset, limit) {
  const url = new URL(`${POLYMARKET_BASE_URL}/markets`);
  url.searchParams.set('tag_id', String(tagId));
  url.searchParams.set('related_tags', 'true');
  url.searchParams.set('active', 'true');
  url.searchParams.set('closed', 'false');
  url.searchParams.set('limit', String(limit));
  url.searchParams.set('offset', String(offset));
  url.searchParams.set('sports_market_types', 'moneyline');
  return url.toString();
}

function isPolymarketNbaMoneylineMarket(market) {
  const slug = String(market?.slug ?? '').trim().toLowerCase();
  const question = String(market?.question ?? '').trim().toLowerCase();
  const sportsMarketType = normalizeOutcomeToken(market?.sportsMarketType);

  if (!slug.startsWith('nba-')) return false;
  if (slug.includes('-spread-') || slug.includes('-total-')) return false;
  if (sportsMarketType && sportsMarketType !== 'moneyline') return false;
  if (!question.includes(' vs')) return false;

  try {
    const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug);
    if (Array.isArray(outcomes) && outcomes.length !== 2) return false;
  } catch {
    return false;
  }

  return true;
}

function isPolymarketFootballMoneylineMarket(market) {
  const question = String(market?.question ?? '').trim().toLowerCase();
  const sportsMarketType = normalizeOutcomeToken(market?.sportsMarketType);

  if (sportsMarketType && sportsMarketType !== 'moneyline') return false;
  if (!question.includes(' vs')) return false;

  try {
    const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', String(market?.slug ?? '').trim() || 'unknown');
    if (!Array.isArray(outcomes)) return false;
    if (outcomes.length < 2 || outcomes.length > 3) return false;
  } catch {
    return false;
  }

  return true;
}

async function fetchPolymarketActiveNbaMarkets() {
  const sportsMetadata = await fetchPolymarketSportsMetadata();
  const discoveredTagId = resolvePolymarketNbaTagId(sportsMetadata);
  const tagIdsToTry = discoveredTagId === 745 ? [745] : [discoveredTagId, 745];
  const limit = 200;

  for (const tagId of tagIdsToTry) {
    const allMarkets = [];
    for (let offset = 0; offset < 1000; offset += limit) {
      const batch = await fetchPolymarketJson(
        buildPolymarketMarketsUrl(tagId, offset, limit),
        `NBA markets 读取（tag_id=${tagId}, offset=${offset}）`
      );
      if (!Array.isArray(batch)) {
        throw new Error(`Polymarket NBA markets 返回异常：期望数组，实际为 ${typeof batch}`);
      }

      allMarkets.push(...batch);
      if (batch.length < limit) break;
    }

    const deduped = new Map();
    for (const market of allMarkets) {
      const slug = String(market?.slug ?? '').trim();
      if (!slug || deduped.has(slug)) continue;
      if (!isPolymarketNbaMoneylineMarket(market)) continue;
      deduped.set(slug, market);
    }

    if (deduped.size > 0) {
      return [...deduped.values()];
    }
  }

  throw new Error('Polymarket NBA markets 为空：未获取到任何可用的 moneyline 市场');
}

async function fetchPolymarketActiveFootballMarkets() {
  const tagId = 100350; // 足球总标签
  const limit = 200;
  const allMarkets = [];

  for (let offset = 0; offset < 20000; offset += limit) {
    const batch = await fetchPolymarketJson(
      buildPolymarketMarketsUrl(tagId, offset, limit),
      `Football markets 读取（tag_id=${tagId}, offset=${offset}）`
    );
    if (!Array.isArray(batch)) {
      throw new Error(`Polymarket Football markets 返回异常：期望数组，实际为 ${typeof batch}`);
    }
    allMarkets.push(...batch);
    if (batch.length < limit) break;
  }

  const deduped = new Map();
  for (const market of allMarkets) {
    const slug = String(market?.slug ?? '').trim();
    if (!slug || deduped.has(slug)) continue;
    if (!isPolymarketFootballMoneylineMarket(market)) continue;
    deduped.set(slug, market);
  }

  if (deduped.size === 0) {
    throw new Error('Polymarket Football markets 为空：未获取到任何可用的 moneyline 市场');
  }

  return [...deduped.values()];
}

function parseDateYMD(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';

  const normalized = raw.replace(/[./]/g, '-');
  const match = normalized.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return '';

  const year = match[1];
  const month = match[2].padStart(2, '0');
  const day = match[3].padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function shiftDateYMD(dateYmd, deltaDays) {
  const m = String(dateYmd ?? '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  const date = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]) + Number(deltaDays || 0)));
  if (Number.isNaN(date.getTime())) return '';
  const y = date.getUTCFullYear();
  const mo = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return `${y}-${mo}-${d}`;
}

function textContainsAnyAlias(text, aliases) {
  const normalized = normalizeOutcomeToken(text);
  if (!normalized) return false;
  for (const alias of aliases) {
    if (alias && normalized.includes(alias)) return true;
  }
  return false;
}

function marketContainsDate(market, dateYmd) {
  if (!dateYmd) return false;
  const fields = [
    market?.slug,
    market?.question,
    market?.endDate,
    market?.startDate,
    market?.endDateIso,
    market?.startDateIso,
    market?.gameStartTime,
    market?.eventStartTime
  ];
  const candidates = [
    dateYmd,
    shiftDateYMD(dateYmd, -1),
    shiftDateYMD(dateYmd, 1)
  ].filter(Boolean);

  return fields.some(item => {
    const text = String(item ?? '');
    return candidates.some(candidate => text.includes(candidate));
  });
}

function isYesNoOutcomes(outcomes) {
  if (!Array.isArray(outcomes) || outcomes.length !== 2) return false;
  const normalized = outcomes.map(item => normalizeOutcomeToken(item));
  return normalized.includes('yes') && normalized.includes('no');
}

function getYesProbabilityFromMarket(market) {
  const slug = String(market?.slug ?? '').trim() || 'unknown';
  const outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(String);
  const prices = parseJsonArrayField(market?.outcomePrices ?? '[]', 'outcomePrices', slug);
  if (!Array.isArray(outcomes) || !Array.isArray(prices) || outcomes.length !== prices.length) return NaN;
  const yesIdx = outcomes.findIndex(item => normalizeOutcomeToken(item) === 'yes');
  if (yesIdx === -1) return NaN;
  return toPercentProbability(prices[yesIdx]);
}

function roundThreeWayProbabilities(homeRaw, awayRaw, drawRaw) {
  const h = Number(homeRaw);
  const a = Number(awayRaw);
  const d = Number(drawRaw);
  if (!Number.isFinite(h) || !Number.isFinite(a) || !Number.isFinite(d)) {
    return { homeWin: NaN, awayWin: NaN, drawWin: NaN };
  }

  const total = h + a + d;
  if (!(total > 0)) {
    return { homeWin: NaN, awayWin: NaN, drawWin: NaN };
  }

  const nh = (h / total) * 100;
  const na = (a / total) * 100;
  const nd = (d / total) * 100;

  let homeWin = Math.max(0, Math.min(100, Math.round(nh)));
  let awayWin = Math.max(0, Math.min(100, Math.round(na)));
  let drawWin = Math.max(0, Math.min(100, Math.round(nd)));
  const sum = homeWin + awayWin + drawWin;
  if (sum !== 100) {
    const deltas = [
      { key: 'home', delta: Math.abs(nh - homeWin) },
      { key: 'away', delta: Math.abs(na - awayWin) },
      { key: 'draw', delta: Math.abs(nd - drawWin) }
    ].sort((x, y) => y.delta - x.delta);
    const adjust = 100 - sum;
    const target = deltas[0]?.key ?? 'draw';
    if (target === 'home') homeWin = Math.max(0, Math.min(100, homeWin + adjust));
    if (target === 'away') awayWin = Math.max(0, Math.min(100, awayWin + adjust));
    if (target === 'draw') drawWin = Math.max(0, Math.min(100, drawWin + adjust));
  }
  return { homeWin, awayWin, drawWin };
}

async function resolveFootballOddsFromEvent(row, market, teamsMap, eventCache, fallbackEventSlug = '') {
  const eventSlug = String(fallbackEventSlug ?? '').trim()
    || String(market?.events?.[0]?.slug ?? '').trim()
    || String(market?.events?.[0]?.ticker ?? '').trim();
  if (!eventSlug) return null;

  let eventData = eventCache.get(eventSlug);
  if (!eventData) {
    eventData = await fetchPolymarketEventBySlug(eventSlug);
    eventCache.set(eventSlug, eventData);
  }

  const markets = Array.isArray(eventData?.markets) ? eventData.markets : [];
  if (markets.length === 0) return null;

  const homeAliases = buildTeamAliasTokens(row.home_team, teamsMap);
  const awayAliases = buildTeamAliasTokens(row.away_team, teamsMap);
  let homeMarket = null;
  let awayMarket = null;
  let drawMarket = null;

  for (const m of markets) {
    const title = String(m?.groupItemTitle ?? '');
    const question = String(m?.question ?? '');
    const combined = `${title} ${question}`.trim();
    const isDraw = normalizeOutcomeToken(combined).includes('draw')
      || normalizeOutcomeToken(combined).includes('tie');
    if (isDraw) {
      drawMarket = m;
      continue;
    }

    const matchesHome = textContainsAnyAlias(combined, homeAliases);
    const matchesAway = textContainsAnyAlias(combined, awayAliases);
    if (matchesHome && !homeMarket) {
      homeMarket = m;
      continue;
    }
    if (matchesAway && !awayMarket) {
      awayMarket = m;
    }
  }

  if (!homeMarket || !awayMarket || !drawMarket) return null;

  const homeProb = getYesProbabilityFromMarket(homeMarket);
  const awayProb = getYesProbabilityFromMarket(awayMarket);
  const drawProb = getYesProbabilityFromMarket(drawMarket);
  const rounded = roundThreeWayProbabilities(homeProb, awayProb, drawProb);
  if (!Number.isFinite(rounded.homeWin) || !Number.isFinite(rounded.awayWin) || !Number.isFinite(rounded.drawWin)) {
    return null;
  }
  return {
    homeWin: rounded.homeWin,
    awayWin: rounded.awayWin,
    drawWin: rounded.drawWin
  };
}

function getTeamEnglishName(teamId, teamsMap) {
  const en = String(teamsMap[teamId]?.en ?? '').trim();
  if (!en) {
    throw new Error(`球队 ${teamId} 缺少英文名，无法自动匹配 Polymarket`);
  }
  return en;
}

function scoreMarketCandidate(market, dateYmd, homeAliases, awayAliases, options = {}) {
  const question = String(market?.question ?? '');
  const slug = String(market?.slug ?? '');
  const sport = String(options.sport ?? 'nba').trim().toLowerCase();
  let outcomes = [];
  try {
    outcomes = parseJsonArrayField(market?.outcomes ?? '[]', 'outcomes', slug).map(String);
  } catch {
    outcomes = [];
  }

  const outcomeHome = findOutcomeIndexByAliases(outcomes, homeAliases) !== -1;
  const outcomeAway = findOutcomeIndexByAliases(outcomes, awayAliases) !== -1;
  const textHome = textContainsAnyAlias(question, homeAliases);
  const textAway = textContainsAnyAlias(question, awayAliases);

  const homeMatched = outcomeHome || textHome;
  const awayMatched = outcomeAway || textAway;
  const dateMatched = marketContainsDate(market, dateYmd);
  const nbaHint = slug.startsWith('nba-') || normalizeOutcomeToken(question).includes('nba');
  const footballHint = normalizeOutcomeToken(question).includes('vs');

  let score = 0;
  if (sport === 'nba' && nbaHint) score += 40;
  if (sport === 'football' && footballHint) score += 20;
  if (homeMatched) score += 40;
  if (awayMatched) score += 40;
  if (dateMatched) score += 40;
  if (outcomeHome && outcomeAway) score += 40;
  if (textHome && textAway) score += 20;
  if (market?.active === true) score += 5;
  score += Math.min(Number(market?.volumeNum ?? market?.volume ?? 0) / 10000, 10);
  const requireDateMatch = sport !== 'football';
  const isAccepted = homeMatched && awayMatched && (requireDateMatch ? dateMatched : true);
  return {
    score: isAccepted ? score : -1,
    previewScore: score,
    reason: isAccepted ? 'ok' : 'team/date mismatch',
    homeMatched,
    awayMatched,
    dateMatched
  };
}

function summarizeMarketForError(market, diagnostics = {}) {
  const slug = String(market?.slug ?? '').trim() || '(no slug)';
  const question = String(market?.question ?? '').trim() || '(no question)';
  const marketDate = [
    market?.endDate,
    market?.startDate,
    market?.endDateIso,
    market?.startDateIso,
    market?.gameStartTime,
    market?.eventStartTime
  ].find(Boolean);

  const parts = [`slug=${slug}`, `question=${question}`];
  if (marketDate) parts.push(`date=${marketDate}`);
  if (market?.sportsMarketType) parts.push(`type=${market.sportsMarketType}`);
  if (Object.keys(diagnostics).length > 0) {
    parts.push(
      `homeMatched=${diagnostics.homeMatched ? 'yes' : 'no'}`,
      `awayMatched=${diagnostics.awayMatched ? 'yes' : 'no'}`,
      `dateMatched=${diagnostics.dateMatched ? 'yes' : 'no'}`,
      `previewScore=${diagnostics.previewScore ?? diagnostics.score ?? 'n/a'}`
    );
  }
  return parts.join(' | ');
}

async function resolvePolymarketMarketByGame(row, teamsMap, activeNbaMarkets, searchCache, options = {}) {
  const inputSlug = normalizePolymarketInputSlug(getRowPolymarketInput(row));
  if (inputSlug) {
    return { slug: inputSlug, market: null };
  }

  const dateYmd = parseDateYMD(row.date);
  if (!dateYmd) {
    throw new Error(`日期格式无法识别（${row.date}），请使用 YYYY-MM-DD`);
  }

  const homeEn = getTeamEnglishName(row.home_team, teamsMap);
  const awayEn = getTeamEnglishName(row.away_team, teamsMap);
  const cacheKey = `${dateYmd}|${homeEn}|${awayEn}`;
  if (searchCache.has(cacheKey)) return searchCache.get(cacheKey);

  const homeAliases = buildTeamAliasTokens(row.home_team, teamsMap);
  const awayAliases = buildTeamAliasTokens(row.away_team, teamsMap);

  let best = null;
  let bestScore = -1;
  const inspected = [];
  for (const market of activeNbaMarkets) {
    const result = scoreMarketCandidate(market, dateYmd, homeAliases, awayAliases, options);
    inspected.push({ market, result });
    const { score } = result;
    if (score > bestScore) {
      bestScore = score;
      best = market;
    }
  }

  if (!best || bestScore < 0) {
    const preview = inspected
      .sort((a, b) => (b.result.previewScore ?? b.result.score ?? -1) - (a.result.previewScore ?? a.result.score ?? -1))
      .slice(0, 3)
      .map(({ market, result }) => `- ${summarizeMarketForError(market, result)}`)
      .join('\n');
    const candidateHint = preview ? `\n候选市场预览：\n${preview}` : '\n候选市场预览：无';
    throw new Error(
      `未找到匹配的 Polymarket 市场：${row.home_team} vs ${row.away_team} (${dateYmd})\n` +
      `请检查表格中的 Polymarket 链接、日期和球队 ID，或手动确认这场比赛在 Polymarket 上是否存在。` +
      candidateHint
    );
  }

  const slug = String(best.slug).trim();
  const resolved = { slug, market: best };
  searchCache.set(cacheKey, resolved);
  return resolved;
}

async function enrichRowsWithPolymarketOdds(gamesRows, teamsMap, options = {}) {
  const sport = String(options.sport ?? 'nba').trim().toLowerCase();
  const strictNotFound = options.strictNotFound !== false;
  const includeDraw = options.includeDraw === true;

  const cache = new Map();
  const searchCache = new Map();
  const eventCache = new Map();
  const activeNbaMarkets = sport === 'football'
    ? await fetchPolymarketActiveFootballMarkets()
    : await fetchPolymarketActiveNbaMarkets();
  const nextRows = [];
  let enrichedCount = 0;
  let autoMatchedCount = 0;

  for (const row of gamesRows) {
    const rawPolymarketInput = getRowPolymarketInput(row);
    const inputSlug = normalizePolymarketInputSlug(rawPolymarketInput);
    const rowForResolve = inputSlug
      ? { ...row, polymarket_url: rawPolymarketInput, polymarket_slug: inputSlug }
      : row;

    let resolved;
    try {
      resolved = await resolvePolymarketMarketByGame(
        rowForResolve,
        teamsMap,
        activeNbaMarkets,
        searchCache,
        { sport }
      );
    } catch (err) {
      if (strictNotFound) throw err;
      warnOnce(
        `polymarket-not-found:${sport}:${row.home_team}:${row.away_team}:${row.date}`,
        `未匹配到 ${row.home_team} vs ${row.away_team}（${row.date}）的 Polymarket 市场，保留原赔率`
      );
      nextRows.push({ ...row });
      continue;
    }

    const { slug, market: resolvedMarket } = resolved;
    if (!inputSlug && slug) autoMatchedCount++;

    let market = resolvedMarket ?? cache.get(slug);
    let fallbackEventSlug = '';
    let skipThisRow = false;
    if (!market) {
      try {
        try {
          market = await fetchPolymarketMarketBySlug(slug);
          cache.set(slug, market);
        } catch (err) {
          const canTryEvent = sport === 'football' && includeDraw && String(err?.message ?? '').includes('HTTP 404');
          if (!canTryEvent) throw err;

          const eventData = await fetchPolymarketEventBySlug(slug);
          eventCache.set(slug, eventData);
          const eventMarkets = Array.isArray(eventData?.markets) ? eventData.markets : [];
          market = eventMarkets[0] ?? null;
          fallbackEventSlug = slug;
          if (!market) {
            throw new Error(`Polymarket 事件无可用盘口（${slug}）`);
          }
        }
      } catch (err) {
        if (strictNotFound) throw err;
        warnOnce(
          `polymarket-slug-failed:${sport}:${slug}`,
          `Polymarket 链接/slug 无法解析（${slug}），保留原赔率`
        );
        nextRows.push({ ...row });
        skipThisRow = true;
      }
    }
    if (skipThisRow) continue;

    const outcomes = parseJsonArrayField(market.outcomes, 'outcomes', slug).map(String);
    const outcomePrices = parseJsonArrayField(market.outcomePrices, 'outcomePrices', slug);
    if (outcomes.length === 0 || outcomes.length !== outcomePrices.length) {
      throw new Error(`Polymarket 数据异常（${slug}）：outcomes 与 outcomePrices 数量不一致`);
    }

    if (sport === 'football' && includeDraw && isYesNoOutcomes(outcomes)) {
      const byEvent = await resolveFootballOddsFromEvent(row, market, teamsMap, eventCache, fallbackEventSlug);
      if (byEvent) {
        const originalPolymarketUrl = String(row.polymarket_url ?? '').trim();
        const originalPolymarketSlug = String(row.polymarket_slug ?? '').trim();
        nextRows.push({
          ...row,
          polymarket_url: originalPolymarketUrl || buildPolymarketEventUrl(slug),
          polymarket_slug: slug,
          home_win: String(byEvent.homeWin),
          away_win: String(byEvent.awayWin),
          draw_win: String(byEvent.drawWin),
          __resolvedPolymarketUrl: buildPolymarketEventUrl(slug),
          __resolvedPolymarketSlug: slug,
          __shouldWritePolymarketUrlBack: !originalPolymarketUrl,
          __shouldWritePolymarketSlugBack: !originalPolymarketSlug,
          __autoFilledFromPolymarket: true
        });
        enrichedCount++;
        continue;
      }
    }

    let homeIndex = -1;
    let awayIndex = -1;

    const homeOutcomeOverride = normalizeOutcomeToken(row.home_outcome);
    const awayOutcomeOverride = normalizeOutcomeToken(row.away_outcome);

    if (homeOutcomeOverride) {
      homeIndex = findOutcomeIndexByAliases(outcomes, new Set([homeOutcomeOverride]));
    }
    if (awayOutcomeOverride) {
      awayIndex = findOutcomeIndexByAliases(outcomes, new Set([awayOutcomeOverride]));
    }

    if (homeIndex === -1) {
      homeIndex = findOutcomeIndexByAliases(outcomes, buildTeamAliasTokens(row.home_team, teamsMap));
    }
    if (awayIndex === -1) {
      awayIndex = findOutcomeIndexByAliases(outcomes, buildTeamAliasTokens(row.away_team, teamsMap));
    }

    // 二元市场兜底：若只识别到一边，另一边默认取剩余 outcome
    if (outcomes.length === 2) {
      if (homeIndex !== -1 && awayIndex === -1) {
        awayIndex = homeIndex === 0 ? 1 : 0;
      } else if (awayIndex !== -1 && homeIndex === -1) {
        homeIndex = awayIndex === 0 ? 1 : 0;
      }
    }

    if (homeIndex === -1 || awayIndex === -1 || homeIndex === awayIndex) {
      throw new Error(
        `Polymarket outcome 映射失败（${slug}）：home=${row.home_team}, away=${row.away_team}, outcomes=${JSON.stringify(outcomes)}`
      );
    }

    const rawHomeWin = toPercentProbability(outcomePrices[homeIndex]);
    const rawAwayWin = toPercentProbability(outcomePrices[awayIndex]);

    let homeWin;
    let awayWin;
    let drawWin = '';

    if (includeDraw) {
      let drawIndex = findOutcomeIndexByAliases(outcomes, new Set(['draw', 'tie']));
      if (drawIndex === -1 && outcomes.length === 3) {
        drawIndex = [0, 1, 2].find(idx => idx !== homeIndex && idx !== awayIndex);
      }

      const rawDrawWin = drawIndex === -1 ? NaN : toPercentProbability(outcomePrices[drawIndex]);
      if (Number.isFinite(rawHomeWin) && Number.isFinite(rawAwayWin) && Number.isFinite(rawDrawWin)) {
        homeWin = Math.max(0, Math.min(100, Math.round(rawHomeWin)));
        awayWin = Math.max(0, Math.min(100, Math.round(rawAwayWin)));
        let drawRate = Math.max(0, Math.min(100, Math.round(rawDrawWin)));
        const sum = homeWin + awayWin + drawRate;
        if (sum !== 100) drawRate = Math.max(0, Math.min(100, drawRate + (100 - sum)));
        drawWin = String(drawRate);
      } else {
        homeWin = Number.isFinite(rawHomeWin) ? Math.max(0, Math.min(100, Math.round(rawHomeWin))) : NaN;
        awayWin = Number.isFinite(rawAwayWin) ? Math.max(0, Math.min(100, Math.round(rawAwayWin))) : NaN;
      }
    } else {
      const rounded = roundWinRatesToIntegers(rawHomeWin, rawAwayWin);
      homeWin = rounded.homeWin;
      awayWin = rounded.awayWin;
    }

    if (!Number.isFinite(homeWin) || !Number.isFinite(awayWin)) {
      throw new Error(`Polymarket 概率解析失败（${slug}）：outcomePrices=${JSON.stringify(outcomePrices)}`);
    }

    const originalPolymarketUrl = String(row.polymarket_url ?? '').trim();
    const originalPolymarketSlug = String(row.polymarket_slug ?? '').trim();
    nextRows.push({
      ...row,
      polymarket_url: originalPolymarketUrl || buildPolymarketEventUrl(slug),
      polymarket_slug: slug,
      home_win: String(homeWin),
      away_win: String(awayWin),
      draw_win: drawWin,
      __resolvedPolymarketUrl: buildPolymarketEventUrl(slug),
      __resolvedPolymarketSlug: slug,
      __shouldWritePolymarketUrlBack: !originalPolymarketUrl,
      __shouldWritePolymarketSlugBack: !originalPolymarketSlug,
      __autoFilledFromPolymarket: true
    });
    enrichedCount++;
  }

  return { rows: nextRows, enrichedCount, autoMatchedCount };
}

// ── 综合事件：从 Lark 读取 ──────────────────────────────────
async function fetchComprehensiveEventsFromLark(config, accessToken, sheetId, sheetLabel = 'comprehensiveSheetId', templateKey = 'comprehensive') {
  const resolvedSheetId = String(sheetId ?? '').trim();
  if (!resolvedSheetId) {
    throw new Error(`lark.config.json 缺少 ${sheetLabel}`);
  }

  const spreadsheetToken = config.spreadsheetToken;
  const range = encodeURIComponent(`${resolvedSheetId}!A1:Z35`);
  const url = `https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 综合事件表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('综合事件表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  const headers = values[0].map(cell => String(cell ?? '').trim().toLowerCase());

  // 优先找第一列为 "en" 的行（用户手动填写的英文源内容），找不到才退化到 row 2
  const enRow = values.slice(1).find(r => String(r?.[0] ?? '').trim().toLowerCase() === 'en');
  const row = (enRow ?? values[1]).map(cell => String(cell ?? '').trim());

  function col(name) {
    const idx = headers.indexOf(name);
    return idx !== -1 ? row[idx] : '';
  }

  // 图片 URL 是语言无关的，始终从第 2 行（原始数据行）读取
  const baseRow = values[1].map(cell => String(cell ?? '').trim());
  function rawCol(name) {
    const idx = headers.indexOf(name);
    return idx !== -1 ? baseRow[idx] : '';
  }

  const MAX_CARDS = 4;
  const cardIndexes = headers
    .map((header) => {
      const m = /^(\d+)$/.exec(header);
      if (!m) return null;
      const index = Number(m[1]);
      if (!Number.isInteger(index) || index <= 0) return null;
      if (!headers.includes(`percent_${index}`)) return null;
      return index;
    })
    .filter((v) => v !== null)
    .sort((a, b) => a - b)
    .slice(0, MAX_CARDS);

  const cards = (cardIndexes.length > 0 ? cardIndexes : [1, 2, 3])
    .map((index) => {
      const percentRaw = Number(col(`percent_${index}`));
      // Lark 存的是小数（0.53 = 53%），需转换为整数百分比
      const percent = Number.isFinite(percentRaw)
        ? (percentRaw > 0 && percentRaw <= 1 ? Math.round(percentRaw * 100) : Math.round(percentRaw))
        : 50;
      return {
        text: sanitizeCardText(col(String(index))),
        percent,
        image: rawCol(`image_${index}`)
      };
    });

  if (templateKey === 'worldcup') {
    const scenarioConfig = parseComprehensiveScenarioConfig(values, 20, 35);
    const scenarioType = normalizeConfigKey(scenarioConfig.scenario_type);
    const marketSlugInput = String(scenarioConfig.market_slug || scenarioConfig.market_url || '').trim();
    const cardCount = Math.max(2, Math.min(4, Number(scenarioConfig.card_count) || cards.length || 3));

    if (scenarioType && !['worldcup_winner', 'group_winner'].includes(scenarioType)) {
      warnOnce(`worldcup-scenario-unknown:${scenarioType}`, `未知 scenario_type：${scenarioType}（支持 worldcup_winner / group_winner）`);
    }

    if (marketSlugInput) {
      const footballTeamsMap = loadTeams(FOOTBALL_TEAMS_CSV);
      const autoCards = await buildWorldCupCardsFromPolymarketInput(marketSlugInput, cardCount, footballTeamsMap);
      const key = extractPolymarketSlug(marketSlugInput) || marketSlugInput;

      if (autoCards.length > 0) {
        cards.splice(0, cards.length, ...autoCards);
        await writeBackComprehensiveCardsToLark(cards, accessToken, spreadsheetToken, resolvedSheetId);
      } else {
        warnOnce(`worldcup-no-cards:${key}`, `Polymarket 输入未解析出可用卡片：${key}`);
      }
    }
  }

  // 读取表格里已有的翻译行（row 3+，A 列为语言代码）
  const existingTranslations = {};
  for (const translatedRow of values.slice(2)) {
    const lang = String(translatedRow?.[0] ?? '').trim().toLowerCase();
    if (!lang) continue;
    const r = translatedRow.map(cell => String(cell ?? '').trim());
    const colByIdx = (idx) => r[idx] ?? '';
    const cardIndexesForTranslation = cardIndexes.length > 0 ? cardIndexes : [1, 2, 3];
    existingTranslations[lang] = {
      mainTitle: colByIdx(headers.indexOf('main title')),
      subTitle: colByIdx(headers.indexOf('sub title')),
      footer: colByIdx(headers.indexOf('footer')),
      cards: cardIndexesForTranslation.map((index, i) => ({
        text: sanitizeCardText(colByIdx(headers.indexOf(String(index)))),
        percent: cards[i]?.percent ?? 50
      }))
    };
  }

  return {
    mainTitle: col('main title'),
    subTitle: col('sub title'),
    footer: col('footer'),
    cards,
    existingTranslations
  };
}

// ── 综合事件：调用 MyMemory 免费翻译（无需 API Key）──────────
async function translateOneText(text, fromLang, toLang, retries = 3) {
  if (!text) return text;
  const params = new URLSearchParams({ q: text, langpair: `${fromLang}|${toLang}` });
  for (let attempt = 0; attempt <= retries; attempt++) {
    const res = await fetch(`https://api.mymemory.translated.net/get?${params}`);
    if (res.status === 429) {
      if (attempt < retries) {
        await new Promise(r => setTimeout(r, 3000 * (attempt + 1)));
        continue;
      }
      throw new Error(`MyMemory 翻译请求失败（${fromLang}→${toLang}）：HTTP ${res.status}`);
    }
    if (!res.ok) {
      throw new Error(`MyMemory 翻译请求失败（${fromLang}→${toLang}）：HTTP ${res.status}`);
    }
    const data = await res.json();
    if (data.responseStatus !== 200) {
      throw new Error(`MyMemory 翻译失败（${fromLang}→${toLang}）：${data.responseDetails}`);
    }
    return String(data.responseData?.translatedText ?? text);
  }
}

async function translateComprehensiveData(sourceData, targetLangs, fromLang = 'zh-CN') {
  // 需要翻译的文本：标题、副标题、footer、N 张卡片问题
  const texts = [
    sourceData.mainTitle,
    sourceData.subTitle,
    sourceData.footer,
    ...sourceData.cards.map(c => c.text)
  ];

  const result = {};

  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(texts.map(t => translateOneText(t, fromLang, lang)));
    result[lang] = {
      mainTitle: translated[0],
      subTitle: translated[1],
      footer: translated[2],
      cards: sourceData.cards.map((card, i) => ({
        text: sanitizeCardText(translated[3 + i]),
        percent: card.percent
      }))
    };
    console.log(' ✅');
  }

  return result;
}

// ── 全球预测市场：从 Lark 读取 title/stat/desc 三列 ───────────
async function fetchGlobalCardsFromLark(config, accessToken, sheetId, sheetLabel = 'globalSheetId') {
  const resolvedSheetId = String(sheetId ?? '').trim();
  if (!resolvedSheetId) {
    throw new Error(`lark.config.json 缺少 ${sheetLabel}`);
  }

  const spreadsheetToken = config.spreadsheetToken;
  const range = encodeURIComponent(`${resolvedSheetId}!A1:Z40`);
  const url = `https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 全球预测市场表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('全球预测市场表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  const headers = values[0].map(cell => String(cell ?? '').trim().toLowerCase());
  const rowValues = values.slice(1);

  function getHeaderIndex(name, fallbackIndex) {
    const index = headers.indexOf(name);
    return index !== -1 ? index : fallbackIndex;
  }

  const titleIndex = getHeaderIndex('title', 1);
  const statIndex = getHeaderIndex('stat', 2);
  const descIndex = getHeaderIndex('desc', 3);

  const cards = rowValues
    .map((row) => {
      const title = String(row?.[titleIndex] ?? '').trim();
      const stat = String(row?.[statIndex] ?? '').trim();
      const desc = String(row?.[descIndex] ?? '').trim();
      return { title, stat, desc };
    })
    .filter(card => card.title || card.stat || card.desc)
    .slice(0, 3);

  if (cards.length === 0) {
    throw new Error('全球预测市场表格未读取到卡片数据，请检查 title/stat/desc 列');
  }

  return cards;
}

// ── 全球预测市场：翻译 title/desc，stat 原样保留 ─────────────
async function translateGlobalCards(sourceCards, targetLangs, fromLang = 'zh-CN') {
  const result = {};
  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(
      sourceCards.flatMap(card => [
        translateOneText(card.title, fromLang, lang),
        translateOneText(card.desc, fromLang, lang)
      ])
    );
    result[lang] = sourceCards.map((card, i) => ({
      title: translated[i * 2],
      stat: String(card.stat ?? ''),
      desc: translated[i * 2 + 1]
    }));
    console.log(' ✅');
  }
  return result;
}

function buildGlobalPosterPayload(cards, templateConfig) {
  const defaultIcons = ['1.png', '2.png', '3.png'];
  const logosDir = String(templateConfig.logosDir ?? '').trim();

  return {
    cards: cards.map((card, index) => {
      const iconFile = defaultIcons[index] || defaultIcons[defaultIcons.length - 1];
      const iconPath = logosDir
        ? path.join(logosDir, iconFile)
        : path.join(BASE_DIR, 'assets', 'logos', '全球预测市场', iconFile);
      return {
        icon: `file://${iconPath}`,
        title: String(card.title ?? '').trim(),
        stat: String(card.stat ?? '').trim(),
        desc: String(card.desc ?? '').trim()
      };
    })
  };
}

// ── 币价预测：从 Lark 读取主文案 + token/price/percent ─────────
async function fetchCoinPriceDataFromLark(config, accessToken, sheetId, sourceLang = 'zh-CN', sheetLabel = 'coinPriceSheetId') {
  const resolvedSheetId = String(sheetId ?? '').trim();
  if (!resolvedSheetId) {
    throw new Error(`lark.config.json 缺少 ${sheetLabel}`);
  }

  const spreadsheetToken = config.spreadsheetToken;
  const range = encodeURIComponent(`${resolvedSheetId}!A1:Z40`);
  const url = `https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 币价预测表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('币价预测表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  const headers = values[0].map(cell => String(cell ?? '').trim().toLowerCase());
  const rows = values.slice(1).map(row => row.map(cell => String(cell ?? '').trim()));

  function getCell(row, name, fallback = '') {
    const idx = headers.indexOf(String(name).toLowerCase());
    return idx !== -1 ? String(row[idx] ?? '').trim() : fallback;
  }

  const cardIndexes = headers
    .map((header) => {
      const m = /^(\d+)$/.exec(header);
      if (!m) return null;
      const index = Number(m[1]);
      if (!Number.isInteger(index) || index <= 0) return null;
      if (!headers.includes(`token_${index}`) || !headers.includes(`percent_${index}`)) return null;
      return index;
    })
    .filter((v) => v !== null)
    .sort((a, b) => a - b)
    .slice(0, 3);

  const sourceLangNormalized = String(sourceLang || 'zh-CN').trim().toLowerCase();
  const sourceRow = rows.find(row => String(row[0] ?? '').trim().toLowerCase() === sourceLangNormalized)
    || rows.find(row => String(row[0] ?? '').trim())
    || rows[0];

  const cards = (cardIndexes.length > 0 ? cardIndexes : [1, 2, 3])
    .map((index) => {
      const token = getCell(sourceRow, `token_${index}`);
      const target = getCell(sourceRow, String(index));
      const percentRaw = Number(getCell(sourceRow, `percent_${index}`));
      const percent = Number.isFinite(percentRaw) ? Math.max(0, Math.min(100, percentRaw)) : 50;
      return {
        token,
        target,
        percent
      };
    })
    .filter(card => card.token || card.target);

  if (cards.length === 0) {
    throw new Error('币价预测表格未读取到卡片数据，请检查 token_i / i / percent_i 列');
  }

  return {
    mainTitle: getCell(sourceRow, 'main title'),
    subTitle: getCell(sourceRow, 'sub title'),
    footer: getCell(sourceRow, 'footer'),
    cards
  };
}

// ── 币价预测：只翻译标题/副标题/footer（卡片只显示币价，无需翻译）──
async function translateCoinPriceData(sourceData, targetLangs, fromLang = 'zh-CN') {
  const texts = [
    sourceData.mainTitle,
    sourceData.subTitle,
    sourceData.footer
  ];

  const result = {};
  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(texts.map(t => translateOneText(t, fromLang, lang)));
    result[lang] = {
      mainTitle: translated[0],
      subTitle: translated[1],
      footer: translated[2],
      cards: sourceData.cards.map((card) => ({
        token: String(card.token ?? ''),
        target: String(card.target ?? ''),
        percent: card.percent
      }))
    };
    console.log(' ✅');
  }

  return result;
}

async function writeBackCoinPriceTranslationsToLark(sourceData, translationsMap, accessToken, spreadsheetToken, sheetId, sourceLang = 'zh-CN') {
  const sourceLangNormalized = String(sourceLang ?? '').trim().toLowerCase();
  const langs = Object.keys(translationsMap)
    .filter(lang => String(lang ?? '').trim().toLowerCase() !== sourceLangNormalized);

  if (langs.length === 0) return 0;

  const MAX_CARDS = 3;
  const totalColumns = 4 + (MAX_CARDS * 3); // A:lang B:title C:subtitle D:footer E~M:token/value/percent
  const blankRow = Array.from({ length: totalColumns }, () => '');

  const rows = langs.map((lang) => {
    const d = translationsMap[lang];
    const row = [
      lang,
      String(d.mainTitle ?? ''),
      String(d.subTitle ?? ''),
      String(d.footer ?? '')
    ];
    for (let i = 0; i < MAX_CARDS; i++) {
      const card = d.cards?.[i];
      row.push(
        String(card?.token ?? ''),
        String(card?.target ?? ''),
        card?.percent ?? ''
      );
    }
    while (row.length < totalColumns) row.push('');
    return row;
  });

  const CLEAR_WINDOW_ROWS = 17; // rows 3..19
  const paddedRows = Array.from({ length: CLEAR_WINDOW_ROWS }, (_, i) => rows[i] ?? blankRow);
  const startRow = 3;
  const endRow = startRow + paddedRows.length - 1;
  const endCol = indexToColumnLabel(totalColumns - 1);
  const range = `${sheetId}!A${startRow}:${endCol}${endRow}`;

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({ valueRange: { range, values: paddedRows } })
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`币价预测翻译回填 Lark 失败：${data.msg || res.status}`);
  }

  return rows.length;
}

// ── 足球赛事：从 Lark 读取主文案 + 3 场比赛 ─────────────────────
// 表头示例：
// lang | title | subtitle | footer | match1_home | match1_away | match1_date | ...
// 可选：match1_link / match2_link / match3_link（优先用链接抓赔率）
async function fetchFootballDataFromLark(config, accessToken, sheetId, sourceLang = 'zh-CN', sheetLabel = 'footballSheetId') {
  const resolvedSheetId = String(sheetId ?? '').trim();
  if (!resolvedSheetId) {
    throw new Error(`lark.config.json 缺少 ${sheetLabel}`);
  }

  const spreadsheetToken = config.spreadsheetToken;
  const range = encodeURIComponent(`${resolvedSheetId}!A1:Q60`);
  const url = `https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}?valueRenderOption=ToString&dateTimeRenderOption=FormattedString`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0) {
    throw new Error(`Lark 足球赛事表格读取失败：${data.msg || res.status}`);
  }

  const values = data.data?.valueRange?.values ?? [];
  if (values.length < 2) {
    throw new Error('足球赛事表格没有可用数据，至少需要 1 行表头和 1 行内容');
  }

  const headers = values[0].map(cell => cellToText(cell));
  const rows = values.slice(1)
    .map(row => row.map(cell => cellToText(cell)))
    .filter(row => row.some(Boolean));

  const sourceLangNormalized = String(sourceLang || 'zh-CN').trim().toLowerCase();
  const sourceRow = rows.find(row => String(row[0] ?? '').trim().toLowerCase() === sourceLangNormalized)
    ?? rows[0]
    ?? [];

  const getCell = (row, name, fallback = '') => {
    const idx = findHeaderIndex(headers, [name]);
    if (idx === -1) return fallback;
    return String(row[idx] ?? '').trim() || fallback;
  };

  function getMatchLink(row, index) {
    const candidates = [
      `match${index}_link`,
      `match${index}_slug`,
      `match${index}_url`,
      `link_${index}`,
      `slug_${index}`,
      `url_${index}`
    ];
    for (const key of candidates) {
      const value = getCell(row, key);
      if (value) return value;
    }
    return '';
  }

  const gamesRows = [];
  for (let i = 1; i <= 3; i++) {
    const homeTeam = getCell(sourceRow, `match${i}_home`);
    const awayTeam = getCell(sourceRow, `match${i}_away`);
    const matchDate = getCell(sourceRow, `match${i}_date`);
    const homeWin = getCell(sourceRow, `match${i}_home_win`);
    const awayWin = getCell(sourceRow, `match${i}_away_win`);

    const matchLink = getMatchLink(sourceRow, i);
    const slug = normalizePolymarketInputSlug(matchLink);

    if (!homeTeam || !awayTeam) {
      if (!slug) continue;
      // link-only row: team info will be fetched from Polymarket
      gamesRows.push({
        date: matchDate,
        home_team: '',
        away_team: '',
        polymarket_slug: slug,
        home_win: '',
        away_win: '',
        link_only: true
      });
      continue;
    }

    gamesRows.push({
      date: matchDate,
      home_team: homeTeam,
      away_team: awayTeam,
      polymarket_slug: slug,
      home_win: String(homeWin).replace('%', '').trim(),
      away_win: String(awayWin).replace('%', '').trim()
    });
  }

  if (gamesRows.length === 0) {
    throw new Error('足球赛事表格没有可用比赛数据（请检查 match1~match3 列或 match1~match3_link 列）');
  }

  return {
    sourceData: {
      mainTitle: getCell(sourceRow, 'title'),
      subTitle: getCell(sourceRow, 'subtitle'),
      footer: getCell(sourceRow, 'footer')
    },
    gamesRows
  };
}

// ── 足球赛事：仅翻译标题/副标题/footer（球队名来自 CSV）──
async function translateFootballTitles(sourceData, targetLangs, fromLang = 'zh-CN') {
  const texts = [sourceData.mainTitle, sourceData.subTitle, sourceData.footer];
  const result = {};

  for (const lang of targetLangs) {
    process.stdout.write(`  → 翻译 ${lang}...`);
    const translated = await Promise.all(texts.map(t => translateOneText(t, fromLang, lang)));
    result[lang] = {
      mainTitle: translated[0],
      subTitle: translated[1],
      footer: translated[2]
    };
    console.log(' ✅');
  }

  return result;
}

// ── 足球赛事：回填翻译结果到 Lark（第 3 行起，A~D）──
async function writeBackFootballTranslationsToLark(sourceData, translationsMap, accessToken, spreadsheetToken, sheetId, sourceLang = 'zh-CN') {
  const sourceLangNormalized = String(sourceLang ?? '').trim().toLowerCase();
  const langs = Object.keys(translationsMap)
    .filter(lang => String(lang ?? '').trim().toLowerCase() !== sourceLangNormalized);

  if (langs.length === 0) return 0;

  const translatedRows = langs.map(lang => {
    const d = translationsMap[lang] ?? {};
    return [
      lang,
      String(d.mainTitle ?? ''),
      String(d.subTitle ?? ''),
      String(d.footer ?? '')
    ];
  });

  const CLEAR_WINDOW_ROWS = 17; // rows 3..19
  const blankRow = ['', '', '', ''];
  const rows = Array.from({ length: CLEAR_WINDOW_ROWS }, (_, i) => translatedRows[i] ?? blankRow);
  const startRow = 3;
  const endRow = startRow + rows.length - 1;
  const range = `${sheetId}!A${startRow}:D${endRow}`;

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({ valueRange: { range, values: rows } })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`足球赛事翻译回填 Lark 失败：${data.msg || res.status}`);
  }

  return translatedRows.length;
}

function buildFootballPosterPayload(gamesRows, teamsMap, lang, sourceData, translationsMap, copyConfig) {
  const templateCopy = copyConfig?.football ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };
  const translated = translationsMap[lang] ?? sourceData ?? {};

  const games = gamesRows.map((row) => {
    const homeId = row.home_team;
    const awayId = row.away_team;

    if (!teamsMap[homeId]) {
      warnOnce(`football-team-missing:${homeId}`, `足球球队 ID 不存在：${homeId}（home_team）`);
    }
    if (!teamsMap[awayId]) {
      warnOnce(`football-team-missing:${awayId}`, `足球球队 ID 不存在：${awayId}（away_team）`);
    }

    const homeRaw = Number(row.home_win);
    const awayRaw = Number(row.away_win);
    const drawRaw = Number(row.draw_win);
    const homeRate = Number.isFinite(homeRaw) ? Math.max(0, Math.min(100, Math.round(homeRaw))) : 0;
    const awayRate = Number.isFinite(awayRaw) ? Math.max(0, Math.min(100, Math.round(awayRaw))) : 0;
    const drawRate = Number.isFinite(drawRaw)
      ? Math.max(0, Math.min(100, Math.round(drawRaw)))
      : ((homeRate + awayRate) > 0 ? Math.max(0, 100 - homeRate - awayRate) : 0);

    return {
      matchDate: row.date,
      homeTeam: {
        name: teamsMap[homeId]?.[lang] ?? teamsMap[homeId]?.en ?? homeId,
        logo: findLogoPath(homeId, teamsMap) ?? '',
        winRate: homeRate
      },
      awayTeam: {
        name: teamsMap[awayId]?.[lang] ?? teamsMap[awayId]?.en ?? awayId,
        logo: findLogoPath(awayId, teamsMap) ?? '',
        winRate: awayRate
      },
      drawRate
    };
  });

  return {
    games,
    copy: {
      title: String(translated.mainTitle ?? mergedCopy.title ?? '').trim(),
      subtitle: String(translated.subTitle ?? mergedCopy.subtitle ?? '').trim(),
      footer: String(translated.footer ?? mergedCopy.footer ?? '').trim(),
      titleFontSize: Number(mergedCopy.titleFontSize ?? 110),
      titleLineHeight: Number(mergedCopy.titleLineHeight ?? 1.2),
      titleMaxWidth: Number(mergedCopy.titleMaxWidth ?? 824),
      subtitleFontSize: Number(mergedCopy.subtitleFontSize ?? 46),
      subtitleLineHeight: Number(mergedCopy.subtitleLineHeight ?? 1.3)
    }
  };
}

function resolveCoinLogoPath(token, templateConfig) {
  const logosDir = String(templateConfig?.logosDir ?? '').trim();
  if (!logosDir) return '';
  const baseToken = String(token ?? '').trim();
  if (!baseToken) return '';

  const normalized = baseToken.replace(/\s+/g, '');
  const candidates = [
    normalized,
    normalized.toUpperCase(),
    normalized.toLowerCase()
  ];
  const exts = ['.png', '.jpg', '.jpeg', '.webp'];
  for (const name of candidates) {
    for (const ext of exts) {
      const absolute = path.join(logosDir, `${name}${ext}`);
      if (fs.existsSync(absolute)) {
        return `file://${absolute}`;
      }
    }
  }
  return '';
}

function formatCoinTargetValue(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';

  const normalized = raw.replace(/,/g, '');
  if (!/^[-+]?\d+(\.\d+)?$/.test(normalized)) {
    return raw;
  }

  const [intPart, decimalPart] = normalized.split('.');
  const sign = intPart.startsWith('-') ? '-' : '';
  const absoluteInt = intPart.replace(/^[-+]/, '');
  const groupedInt = absoluteInt.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return decimalPart ? `${sign}${groupedInt}.${decimalPart}` : `${sign}${groupedInt}`;
}

function buildCoinPricePosterPayload(sourceData, translationsMap, lang, templateConfig) {
  const translated = translationsMap[lang] ?? sourceData;
  const outcomeLabel = 'Yes';

  return {
    games: (translated.cards ?? sourceData.cards).map((card) => {
      const formattedTarget = formatCoinTargetValue(card.target);
      const logo = resolveCoinLogoPath(card.token, templateConfig);
      return {
        date: '',
        homeTeam: {
          name: String(card.token ?? ''),
          logo,
          winRate: Number(card.percent ?? 50)
        },
        awayTeam: {
          name: formattedTarget,
          logo: '',
          winRate: Math.max(0, 100 - Number(card.percent ?? 50))
        }
      };
    }),
    copy: {
      title: String(translated.mainTitle ?? '').trim(),
      subtitle: String(translated.subTitle ?? '').trim(),
      footer: String(translated.footer ?? '').trim(),
      outcomeLabel,
      titleFontSize: 110,
      titleLineHeight: 1.2,
      titleMaxWidth: 824,
      subtitleFontSize: 46,
      subtitleLineHeight: 1.3,
      cardTextFontSize: 60,
      cardTextLineHeight: 1.2,
      cards: (translated.cards ?? sourceData.cards).map((card) => ({
        text: formatCoinTargetValue(card.target),
        image: resolveCoinLogoPath(card.token, templateConfig),
        valueLabel: outcomeLabel
      }))
    }
  };
}

// ── 综合事件：翻译结果回填到 Lark 表格（第 3 行起，仅回填非源语言）──
async function writeBackTranslationsToLark(
  sourceData,
  translationsMap,
  accessToken,
  spreadsheetToken,
  sheetId,
  sourceLang = 'zh-CN',
  { includeSourceLang = false } = {}
) {
  const sourceLangNormalized = String(sourceLang ?? '').trim().toLowerCase();
  // includeSourceLang=true 时把 source lang 也写回（保持其在回填区域的固定位置）
  const langs = Object.keys(translationsMap)
    .filter(lang => includeSourceLang || String(lang ?? '').trim().toLowerCase() !== sourceLangNormalized)
    .sort();

  const MAX_CARDS = 4;
  const cardCount = Math.max(1, Math.min(MAX_CARDS, Array.isArray(sourceData.cards) ? sourceData.cards.length : 0));
  const totalColumns = 4 + (MAX_CARDS * 2); // A:lang B:title C:subtitle D:footer E~L:cards
  const blankRow = Array.from({ length: totalColumns }, () => '');

  const translatedRows = langs.map(lang => {
    const d = translationsMap[lang];
    const row = [
      lang,
      String(d.mainTitle ?? ''),
      String(d.subTitle ?? ''),
      String(d.footer ?? '')
    ];

    for (let i = 0; i < cardCount; i++) {
      row.push(
        String(d.cards?.[i]?.text ?? ''),
        sourceData.cards?.[i]?.percent ?? ''
      );
    }
    while (row.length < totalColumns) row.push('');
    return row;
  });

  if (translatedRows.length === 0) return 0;

  // 预留 A20 开始的配置区（scenario_type / market_slug / card_count），
  // 翻译回填只覆盖 A3~L19，避免清空用户配置。
  const MAX_TRANSLATION_ROWS = 17; // rows 3..19
  if (translatedRows.length > MAX_TRANSLATION_ROWS) {
    throw new Error(`翻译语种过多（${translatedRows.length}），超过表格预留区域上限 ${MAX_TRANSLATION_ROWS}`);
  }

  // 额外填充空行，清掉旧残留回填内容
  const CLEAR_WINDOW_ROWS = MAX_TRANSLATION_ROWS;
  const rows = Array.from({ length: CLEAR_WINDOW_ROWS }, (_, i) => translatedRows[i] ?? blankRow);

  const startRow = 3;
  const endRow = startRow + rows.length - 1;
  const endCol = indexToColumnLabel(totalColumns - 1);
  const range = `${sheetId}!A${startRow}:${endCol}${endRow}`;

  const res = await fetch(`https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify({ valueRange: { range, values: rows } })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.code !== 0) {
    throw new Error(`翻译回填 Lark 失败：${data.msg || res.status}`);
  }
  return translatedRows.length;
}

// ── 综合事件：下载并裁剪卡片图片 URL → 本地临时文件 ────────────
async function downloadAndCropCardImages(cards) {
  const sharp = require('sharp');
  const tmpDir = path.join(BASE_DIR, '_tmp_card_images');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });

  await Promise.all(cards.map(async (card, i) => {
    const url = String(card.image || '').trim();
    if (!url.startsWith('http://') && !url.startsWith('https://')) return;

    const ext = url.split('?')[0].split('.').pop().split('/').pop().toLowerCase();
    const filename = `card_${i}_${Date.now()}.jpg`;
    const destPath = path.join(tmpDir, filename);

    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buffer = Buffer.from(await res.arrayBuffer());
      await sharp(buffer)
        .resize(185, 189, { fit: 'cover', position: 'centre' })
        .jpeg({ quality: 90 })
        .toFile(destPath);
      card.resolvedImage = `file://${destPath}`;
    } catch (err) {
      console.warn(`  ⚠️  卡片 ${i + 1} 图片下载失败（${url}）：${err.message}`);
    }
  }));
}

// ── 综合事件：构建海报 payload ──────────────────────────────
function buildComprehensivePosterPayload(sourceData, translationsMap, lang, copyConfig, templateKey = 'comprehensive') {
  const templateCopy = copyConfig?.[templateKey] ?? copyConfig?.comprehensive ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };

  // Use translation for this lang; fall back to source (zh-CN) data
  const translated = translationsMap[lang] ?? sourceData;

  const COMPREHENSIVE_LOGO_DIR = path.join(BASE_DIR, 'assets', 'logos', '综合事件');
  const WORLD_CUP_LOGO_DIR = path.join(BASE_DIR, 'assets', 'logos', '足球赛事', '世界杯');
  const SUPPORTED_EXTS = ['.png', '.jpg', '.jpeg', '.webp'];
  const worldCupLogoFiles = (templateKey === 'worldcup' && fs.existsSync(WORLD_CUP_LOGO_DIR))
    ? fs.readdirSync(WORLD_CUP_LOGO_DIR)
      .filter(file => SUPPORTED_EXTS.includes(path.extname(file).toLowerCase()))
      .sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'))
    : [];

  function normalizeText(value) {
    return String(value ?? '').toLowerCase().replace(/[\s\-_:：，。、“”"'`·/\\()（）[\]{}!?]+/g, '');
  }

  function resolveWorldCupLogoPath(cardText = '', index = 0) {
    if (!worldCupLogoFiles.length) return '';

    const normalizedCardText = normalizeText(cardText);
    const match = worldCupLogoFiles.find(file => {
      const name = path.parse(file).name.replace(/^世界杯[-_]?/, '');
      const normalizedName = normalizeText(name);
      return normalizedName && normalizedCardText.includes(normalizedName);
    });
    if (match) return `file://${path.join(WORLD_CUP_LOGO_DIR, match)}`;
    const fallback = worldCupLogoFiles[index % worldCupLogoFiles.length];
    return fallback ? `file://${path.join(WORLD_CUP_LOGO_DIR, fallback)}` : '';
  }

  function resolveComprehensiveLogoPath(index = 0) {
    if (templateKey !== 'comprehensive') return '';
    if (!fs.existsSync(COMPREHENSIVE_LOGO_DIR)) return '';

    const slot = String(index + 1);
    for (const ext of SUPPORTED_EXTS) {
      const candidate = path.join(COMPREHENSIVE_LOGO_DIR, `${slot}${ext}`);
      if (fs.existsSync(candidate)) {
        return `file://${candidate}`;
      }
    }
    return '';
  }

  return {
    games: sourceData.cards.map(card => ({
      date: '',
      homeTeam: { name: '', logo: '', winRate: card.percent },
      awayTeam: { name: '', logo: '', winRate: 100 - card.percent }
    })),
    copy: {
      title: String(translated.mainTitle ?? '').trim(),
      subtitle: String(translated.subTitle ?? '').trim(),
      footer: String(translated.footer ?? '').trim(),
      outcomeLabel: String(mergedCopy.outcomeLabel ?? 'Yes'),
      titleFontSize: Number(mergedCopy.titleFontSize ?? 110),
      titleLineHeight: Number(mergedCopy.titleLineHeight ?? 1.2),
      titleMaxWidth: Number(mergedCopy.titleMaxWidth ?? 824),
      subtitleFontSize: Number(mergedCopy.subtitleFontSize ?? 46),
      subtitleLineHeight: Number(mergedCopy.subtitleLineHeight ?? 1.3),
      cardTextFontSize: templateKey === 'worldcup'
        ? 60
        : Number(mergedCopy.cardTextFontSize ?? 40),
      cardTextLineHeight: Number(mergedCopy.cardTextLineHeight ?? 1.2),
      cards: (translated.cards ?? sourceData.cards).map((card, i) => ({
        text: sanitizeCardText(card.text),
        image: templateKey === 'worldcup'
          ? (String(sourceData.cards?.[i]?.image ?? '').trim()
            || resolveWorldCupLogoPath(sourceData.cards?.[i]?.text ?? card.text, i))
          : (String(sourceData.cards?.[i]?.resolvedImage ?? '').trim() || resolveComprehensiveLogoPath(i) || String(defaultCopy.cards?.[i]?.image ?? '').trim()),
        valueLabel: String(mergedCopy.outcomeLabel ?? 'Yes')
      }))
    }
  };
}

// ── 生成单张海报 ──────────────────────────────────────────
async function generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath, options = {}) {
  const viewportWidth = Number(options.viewportWidth ?? 1200);
  const viewportHeight = Number(options.viewportHeight ?? 1200);
  const outputWidth = Number(options.outputWidth ?? viewportWidth);
  const outputHeight = Number(options.outputHeight ?? viewportHeight);

  if (options.viewportWidth || options.viewportHeight) {
    await page.setViewport({
      width: viewportWidth,
      height: viewportHeight,
      deviceScaleFactor: Number(options.deviceScaleFactor ?? 2)
    });
  }

  const injection = `<script>window.GAMES_DATA = ${JSON.stringify(posterPayload)}; window.BG_PATH = "file://${bgPath}";</script>`;
  let html = htmlTemplate.replace('</head>', injection + '\n</head>');

  const tmpFile = path.join(BASE_DIR, '_tmp_poster.html');
  fs.writeFileSync(tmpFile, html, 'utf8');

  await page.goto(`file://${tmpFile}`, { waitUntil: 'networkidle0' });

  await page.evaluate(() => Promise.all([
    document.fonts.ready,
    ...Array.from(document.images).map(img =>
      img.complete ? Promise.resolve() : new Promise(r => { img.onload = r; img.onerror = r; })
    )
  ]));

  await new Promise(r => setTimeout(r, 300));

  const tmpPng = outputPath.replace(/\.(png|jpg|jpeg)$/i, '') + '_raw.png';
  await page.screenshot({
    path: tmpPng,
    clip: { x: 0, y: 0, width: viewportWidth, height: viewportHeight }
  });

  fs.unlinkSync(tmpFile);

  let compressed;
  try {
    compressed = await compressToJpegWithSharp(tmpPng, outputPath, MAX_SIZE_KB, { width: outputWidth, height: outputHeight });
  } finally {
    if (fs.existsSync(tmpPng)) fs.unlinkSync(tmpPng);
  }

  const { sizeKB, profile } = compressed;
  console.log(`  ✅ ${path.basename(outputPath)} (jpg=${sizeKB}KB, profile=${profile}, size=${outputWidth}x${outputHeight})`);
}

// ── 公共：扫描背景图、创建输出目录、启动浏览器 ──────────────
function scanBgFiles(bgDir = BG_DIR) {
  const bgFiles = fs.readdirSync(bgDir).filter(f => f.endsWith('.png') || f.endsWith('.jpg'));
  if (bgFiles.length === 0) {
    console.error(`❌ ${path.basename(bgDir)}/ 目录下没有找到背景图，请先添加语种背景图（如 zh-CN.png）`);
    process.exit(1);
  }
  return bgFiles;
}

function prepareOutputDir(date, subDir) {
  const subDirPath = path.join(OUTPUT_DIR, subDir);
  const dateDir = path.join(subDirPath, date);
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);
  if (!fs.existsSync(subDirPath)) fs.mkdirSync(subDirPath);
  if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir);
  return dateDir;
}

async function launchBrowser() {
  const macChrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
  // 本机有系统 Chrome（Mac 开发用）就走它；否则交给 Puppeteer 自带 Chromium（Linux 服务器场景）
  const opts = {
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--allow-file-access-from-files']
  };
  if (process.platform === 'darwin' && fs.existsSync(macChrome)) {
    opts.executablePath = macChrome;
  }
  return puppeteer.launch(opts);
}

function buildZip(dateDir, outputPrefixWithDate, bgFiles) {
  const zipName = `${outputPrefixWithDate}.zip`;
  const zipPath = path.join(dateDir, zipName);
  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
  const files = bgFiles.map(f => `"${outputPrefixWithDate}_${path.parse(f).name}.jpg"`).join(' ');
  execSync(`cd "${dateDir}" && zip "${zipName}" ${files}`);
  const zipKB = Math.round(fs.statSync(zipPath).size / 1024);
  console.log(`\n📦 ${zipName} (${zipKB}KB)`);
  console.log(`所有图片已保存到：${dateDir}\n`);
}

// ── 主流程 ────────────────────────────────────────────────
async function main() {
  const { templateKey, templateConfig } = parseCliOptions();
  const copyConfig = loadPosterCopyConfig();
  console.log(`\n当前模板：${templateKey} (${path.basename(templateConfig.file)})`);

  // ── F1 流程 ───────────────────────────────────────────────
  if (templateKey === 'f1') {
    console.log('\n从 Lark F1 表格拉取数据...');
    const larkConfig = loadLarkConfig();
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const sourceData = await fetchF1EventsFromLark(larkConfig, accessToken);
    console.log('  ✅ 已读取 F1 事件数据');

    const f1TeamsMap = loadTeams(F1_TEAMS_CSV);

    // ── 若填了 Polymarket URL，自动抓取前3名概率覆盖手动数据 ──
    const polyUrl = String(sourceData.polymarket_url ?? '').trim();
    if (polyUrl) {
      console.log(`\n检测到 Polymarket URL，正在拉取数据：${polyUrl}`);
      try {
        const polyTeams = await fetchF1TeamsFromPolymarket(polyUrl, f1TeamsMap);
        if (polyTeams && polyTeams.length > 0) {
          sourceData.teams = polyTeams;
          console.log(`  ✅ Polymarket 数据已加载（${polyTeams.length} 支车队/车手）`);
          console.log(`     ${polyTeams.map(t => `${String(t._rawName ?? t.teamId).split(/\s+/)[0]}(${t.percent}%)`).join(', ')}`);
        } else {
          console.warn('  ⚠️  Polymarket 返回数据为空，使用 Lark 手动数据');
          console.log(`     车队：${sourceData.teams.map(t => `${t.teamId}(${t.percent}%)`).join(', ')}`);
        }
      } catch (err) {
        console.warn(`  ⚠️  Polymarket 拉取失败（${err.message}），使用 Lark 手动数据`);
        console.log(`     车队：${sourceData.teams.map(t => `${t.teamId}(${t.percent}%)`).join(', ')}`);
      }
    } else {
      console.log(`     车队：${sourceData.teams.map(t => `${t.teamId}(${t.percent}%)`).join(', ')}`);
    }

    // ── 背景图：优先用语言命名文件（zh-CN.png），否则用通用背景（BG.png）──
    const F1_BG_DIR = path.join(BASE_DIR, 'backgrounds-F1');
    const KNOWN_LANGS = new Set(['zh-CN', 'zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi']);
    // 没有语言特定背景时，这些语言都生成（与 Lark 表行对齐）
    const F1_ALL_TARGET_LANGS = ['zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi'];

    const langBgMap = {};   // { lang: absolutePath }
    let genericBgPath = '';

    if (fs.existsSync(F1_BG_DIR)) {
      const bgFileNames = fs.readdirSync(F1_BG_DIR).filter(f => /\.(png|jpg)$/i.test(f));
      for (const f of bgFileNames) {
        const lang = path.parse(f).name;
        if (KNOWN_LANGS.has(lang)) {
          langBgMap[lang] = path.join(F1_BG_DIR, f);
        } else if (!genericBgPath) {
          genericBgPath = path.join(F1_BG_DIR, f);
        }
      }
    }
    if (!genericBgPath && Object.keys(langBgMap).length === 0) {
      console.warn('⚠️  backgrounds-F1/ 目录下没有背景图，将使用纯黑背景');
    }

    // ── 目标语种：有语言背景则由背景文件决定；否则全量翻译 ──
    const sourceLang = larkConfig.sourceLang || 'zh-CN';
    const hasLangBg = Object.keys(langBgMap).length > 0;
    const targetLangs = hasLangBg
      ? Object.keys(langBgMap).filter(l => l !== sourceLang)
      : F1_ALL_TARGET_LANGS;

    let translationsMap = { [sourceLang]: sourceData };

    if (targetLangs.length > 0) {
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')}）...`);
      const translated = await translateF1Titles(sourceData, targetLangs, sourceLang);
      translationsMap = { ...translationsMap, ...translated };

      console.log('\n回填翻译结果到 Lark 表格...');
      const f1Token = larkConfig.f1SpreadsheetToken || larkConfig.spreadsheetToken;
      const writtenCount = await writeBackF1TranslationsToLark(
        sourceData, translationsMap,
        accessToken, f1Token, larkConfig.f1SheetId
      );
      console.log(`  ✅ 已回填 ${writtenCount} 种语言`);
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const langsToGenerate = [sourceLang, ...targetLangs];

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${langsToGenerate.length} 个语种）：`);
    for (const lang of langsToGenerate) {
      const bgPath = hasLangBg
        ? (langBgMap[lang] || genericBgPath || '')
        : (genericBgPath || '');
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.jpg`);
      const posterPayload = buildF1PosterPayload(sourceData, translationsMap, f1TeamsMap, lang, copyConfig);
      await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, langsToGenerate.map(l => `${l}.jpg`));
    return;
  }

  // ── F1 车手流程 ──────────────────────────────────────────
  if (templateKey === 'f1driver') {
    console.log('\n从 Lark F1车手表格拉取数据...');
    const larkConfig = loadLarkConfig();
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const sourceData = await fetchF1DriverEventsFromLark(larkConfig, accessToken);
    console.log('  ✅ 已读取 F1 车手事件数据');

    const driversMap = loadTeams(F1_DRIVER_CSV);

    // ── 若填了 Polymarket URL，自动抓取前4名概率覆盖手动数据 ──
    const polyUrl = String(sourceData.polymarket_url ?? '').trim();
    if (polyUrl) {
      console.log(`\n检测到 Polymarket URL，正在拉取数据：${polyUrl}`);
      try {
        const polyDrivers = await fetchF1DriversFromPolymarket(polyUrl, driversMap);
        if (polyDrivers && polyDrivers.length > 0) {
          sourceData.drivers = polyDrivers;
          console.log(`  ✅ Polymarket 数据已加载（${polyDrivers.length} 位车手）`);
          console.log(`     ${polyDrivers.map(d => `${String(d._rawName ?? d.driverId).split(/\s+/).slice(-1)[0]}(${d.percent}%)`).join(', ')}`);
        } else {
          console.warn('  ⚠️  Polymarket 返回数据为空，使用 Lark 手动数据');
          console.log(`     车手：${sourceData.drivers.map(d => `${d.driverId}(${d.percent}%)`).join(', ')}`);
        }
      } catch (err) {
        console.warn(`  ⚠️  Polymarket 拉取失败（${err.message}），使用 Lark 手动数据`);
        console.log(`     车手：${sourceData.drivers.map(d => `${d.driverId}(${d.percent}%)`).join(', ')}`);
      }
    } else {
      console.log(`     车手：${sourceData.drivers.map(d => `${d.driverId}(${d.percent}%)`).join(', ')}`);
    }

    const F1_DRIVER_BG_DIR = path.join(BASE_DIR, 'backgrounds-F1');
    const KNOWN_LANGS = new Set(['zh-CN', 'zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi']);
    const F1_DRIVER_ALL_TARGET_LANGS = ['zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi'];

    const langBgMap = {};
    let genericBgPath = '';

    if (fs.existsSync(F1_DRIVER_BG_DIR)) {
      const bgFileNames = fs.readdirSync(F1_DRIVER_BG_DIR).filter(f => /\.(png|jpg)$/i.test(f));
      for (const f of bgFileNames) {
        const lang = path.parse(f).name;
        if (KNOWN_LANGS.has(lang)) {
          langBgMap[lang] = path.join(F1_DRIVER_BG_DIR, f);
        } else if (!genericBgPath) {
          genericBgPath = path.join(F1_DRIVER_BG_DIR, f);
        }
      }
    }
    if (!genericBgPath && Object.keys(langBgMap).length === 0) {
      console.warn('⚠️  backgrounds-F1/ 目录下没有背景图，将使用纯黑背景');
    }

    const sourceLang = larkConfig.sourceLang || 'zh-CN';
    const hasLangBg = Object.keys(langBgMap).length > 0;
    const targetLangs = hasLangBg
      ? Object.keys(langBgMap).filter(l => l !== sourceLang)
      : F1_DRIVER_ALL_TARGET_LANGS;

    let translationsMap = { [sourceLang]: sourceData };

    if (targetLangs.length > 0) {
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')}）...`);
      const translated = await translateF1DriverTitles(sourceData, targetLangs, sourceLang);
      translationsMap = { ...translationsMap, ...translated };

      console.log('\n回填翻译结果到 Lark 表格...');
      const driverToken = larkConfig.f1DriverSpreadsheetToken || larkConfig.f1SpreadsheetToken || larkConfig.spreadsheetToken;
      const writtenCount = await writeBackF1DriverTranslationsToLark(
        sourceData, translationsMap,
        accessToken, driverToken, larkConfig.f1DriverSheetId
      );
      console.log(`  ✅ 已回填 ${writtenCount} 种语言`);
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const langsToGenerate = [sourceLang, ...targetLangs];

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${langsToGenerate.length} 个语种）：`);
    for (const lang of langsToGenerate) {
      const bgPath = hasLangBg
        ? (langBgMap[lang] || genericBgPath || '')
        : (genericBgPath || '');
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.jpg`);
      const posterPayload = buildF1DriverPosterPayload(sourceData, translationsMap, driversMap, lang, copyConfig);
      await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, langsToGenerate.map(l => `${l}.jpg`));
    return;
  }

  // ── 足球赛事流程（Lark 文案 + 搜赔率 + 自动翻译回填）──
  if (templateKey === 'football') {
    const larkConfig = loadLarkConfig();
    const sourceLang = larkConfig.sourceLang || 'zh-CN';
    const sheetId = String(larkConfig.footballSheetId || templateConfig.larkSheet || '').trim();

    console.log('\n从 Lark 足球赛事表格拉取数据...');
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const { sourceData, gamesRows: rawGamesRows } = await fetchFootballDataFromLark(
      larkConfig,
      accessToken,
      sheetId,
      sourceLang,
      'footballSheetId'
    );
    console.log(`  ✅ 已读取足球赛事数据（${rawGamesRows.length} 场）`);

    const teamsMap = loadTeams(templateConfig.teamsCsv || FOOTBALL_TEAMS_CSV);

    const linkOnlyCount = rawGamesRows.filter(r => r.link_only).length;
    let resolvedGamesRows = rawGamesRows;
    if (linkOnlyCount > 0) {
      console.log(`\n从 Polymarket 链接获取 ${linkOnlyCount} 场比赛数据...`);
      resolvedGamesRows = await resolveFootballLinkOnlyRows(rawGamesRows, teamsMap);
    }

    let gamesRows = normalizeGameRowsTeamIds(resolvedGamesRows, teamsMap);
    const { rows: enrichedRows, enrichedCount, autoMatchedCount } = await enrichRowsWithPolymarketOdds(
      gamesRows,
      teamsMap,
      { sport: 'football', strictNotFound: false, includeDraw: true }
    );
    gamesRows = enrichedRows;

    if (enrichedCount > 0) {
      console.log(`  ✅ 已从 Polymarket 自动更新 ${enrichedCount} 场赔率`);
    }
    if (autoMatchedCount > 0) {
      console.log(`  ✅ 其中 ${autoMatchedCount} 场由主客队+日期自动匹配到 Polymarket 市场`);
    }

    const bgDir = templateConfig.bgDir || WORLD_CUP_BG_DIR;
    const FOOTBALL_KNOWN_LANGS = new Set(['zh-CN', 'zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi']);
    const FOOTBALL_ALL_TARGET_LANGS = ['zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi'];

    const footballLangBgMap = {};
    let footballGenericBgPath = '';
    const footballBgFileNames = scanBgFiles(bgDir);
    for (const f of footballBgFileNames) {
      const lang = path.parse(f).name;
      if (FOOTBALL_KNOWN_LANGS.has(lang)) {
        footballLangBgMap[lang] = path.join(bgDir, f);
      } else if (!footballGenericBgPath) {
        footballGenericBgPath = path.join(bgDir, f);
      }
    }
    const footballHasLangBg = Object.keys(footballLangBgMap).length > 0;
    const targetLangs = footballHasLangBg
      ? Object.keys(footballLangBgMap).filter(l => l !== sourceLang)
      : FOOTBALL_ALL_TARGET_LANGS;
    const langsToGenerate = [sourceLang, ...targetLangs];

    let translationsMap = { [sourceLang]: sourceData };
    if (targetLangs.length > 0) {
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')}）...`);
      const translated = await translateFootballTitles(sourceData, targetLangs, sourceLang);
      translationsMap = { ...translationsMap, ...translated };

      console.log('\n回填翻译结果到 Lark 表格...');
      const writtenCount = await writeBackFootballTranslationsToLark(
        sourceData,
        translationsMap,
        accessToken,
        larkConfig.spreadsheetToken,
        sheetId,
        sourceLang
      );
      console.log(`  ✅ 已回填 ${writtenCount} 种语言（第 3 行起，A 列为语言代码）`);
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${langsToGenerate.length} 个语种）：`);
    for (const lang of langsToGenerate) {
      const bgPath = footballHasLangBg
        ? (footballLangBgMap[lang] || footballGenericBgPath || '')
        : (footballGenericBgPath || '');
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.jpg`);
      const posterPayload = buildFootballPosterPayload(gamesRows, teamsMap, lang, sourceData, translationsMap, copyConfig);
      await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, langsToGenerate.map(lang => `${lang}.jpg`));
    return;
  }

  // ── 全球预测市场流程（Lark 文案 + 自动翻译 + 多语种背景）──
  if (templateKey === 'coinprice') {
    const larkConfig = loadLarkConfig();
    const sourceLang = larkConfig.sourceLang || 'zh-CN';
    const sheetId = String(larkConfig.coinPriceSheetId || templateConfig.larkSheet || '').trim();

    console.log('\n从 Lark 币价预测表格拉取数据...');
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const sourceData = await fetchCoinPriceDataFromLark(larkConfig, accessToken, sheetId, sourceLang, 'coinPriceSheetId');
    console.log(`  ✅ 已读取币价预测卡片数据（${sourceData.cards.length} 张）`);

    const bgDir = templateConfig.bgDir || GLOBAL_BG_DIR;
    const bgFiles = scanBgFiles(bgDir);
    const targetLangs = bgFiles.map(f => path.parse(f).name).filter(l => l !== sourceLang);

    let translationsMap = { [sourceLang]: sourceData };
    if (targetLangs.length > 0) {
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')})...`);
      const translated = await translateCoinPriceData(sourceData, targetLangs, sourceLang);
      translationsMap = { ...translationsMap, ...translated };

      console.log('\n回填翻译结果到 Lark 表格...');
      const writtenCount = await writeBackCoinPriceTranslationsToLark(
        sourceData,
        translationsMap,
        accessToken,
        larkConfig.spreadsheetToken,
        sheetId,
        sourceLang
      );
      console.log(`  ✅ 已回填 ${writtenCount} 种语言`);
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${bgFiles.length} 个语种）：`);
    for (const bgFile of bgFiles) {
      const lang = path.parse(bgFile).name;
      const bgPath = path.join(bgDir, bgFile);
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.jpg`);
      const payload = buildCoinPricePosterPayload(sourceData, translationsMap, lang, templateConfig);
      await generatePoster(page, htmlTemplate, payload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, bgFiles);
    return;
  }

  // ── 全球预测市场流程（Lark 文案 + 自动翻译 + 多语种背景）──
  if (templateKey === 'global') {
    const larkConfig = loadLarkConfig();
    const sourceLang = larkConfig.sourceLang || 'zh-CN';
    const sheetId = String(larkConfig.globalSheetId || templateConfig.larkSheet || '').trim();

    console.log('\n从 Lark 全球预测市场表格拉取数据...');
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const sourceCards = await fetchGlobalCardsFromLark(larkConfig, accessToken, sheetId, 'globalSheetId');
    console.log(`  ✅ 已读取全球预测市场卡片数据（${sourceCards.length} 张）`);

    const bgDir = templateConfig.bgDir || GLOBAL_BG_DIR;
    const bgFiles = scanBgFiles(bgDir);
    const targetLangs = bgFiles.map(f => path.parse(f).name).filter(l => l !== sourceLang);

    let cardsByLang = { [sourceLang]: sourceCards };
    if (targetLangs.length > 0) {
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')}）...`);
      const translated = await translateGlobalCards(sourceCards, targetLangs, sourceLang);
      cardsByLang = { ...cardsByLang, ...translated };
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${bgFiles.length} 个语种）：`);
    for (const bgFile of bgFiles) {
      const lang = path.parse(bgFile).name;
      const bgPath = path.join(bgDir, bgFile);
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.jpg`);
      const cards = cardsByLang[lang] || sourceCards;
      const payload = buildGlobalPosterPayload(cards, templateConfig);
      await generatePoster(page, htmlTemplate, payload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, bgFiles);
    return;
  }

  // ── 综合事件/世界杯流程 ───────────────────────────────────
  if (templateKey === 'comprehensive' || templateKey === 'worldcup') {
    const larkConfig = loadLarkConfig();
    const templateLabel = templateKey === 'worldcup' ? '世界杯' : '综合事件';
    const sheetId = templateKey === 'worldcup'
      ? String(larkConfig.worldCupSheetId || larkConfig.comprehensiveSheetId || '').trim()
      : String(larkConfig.comprehensiveSheetId || '').trim();
    const sheetFieldLabel = templateKey === 'worldcup'
      ? 'worldCupSheetId（或 comprehensiveSheetId）'
      : 'comprehensiveSheetId';

    console.log(`\n从 Lark ${templateLabel}表格拉取数据...`);
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const sourceData = await fetchComprehensiveEventsFromLark(larkConfig, accessToken, sheetId, sheetFieldLabel, templateKey);
    console.log(`  ✅ 已读取${templateLabel}数据`);

    if (sourceData.cards.some(c => String(c.image || '').startsWith('http'))) {
      console.log('\n下载并裁剪卡片图片...');
      await downloadAndCropCardImages(sourceData.cards);
      console.log('  ✅ 卡片图片处理完成');
    }

    const bgDir = templateConfig.bgDir || BG_DIR;
    const rawBgFiles = scanBgFiles(bgDir);
    const bgFiles = rawBgFiles.filter((bgFile) => {
      const lang = path.parse(bgFile).name.toLowerCase();
      if (templateKey === 'worldcup' && lang === 'bg') return false;
      return true;
    });
    if (bgFiles.length === 0) {
      throw new Error(`模板 ${templateKey} 没有可用语种背景图`);
    }
    // 综合事件使用专属 sourceLang（comprehensiveSourceLang），默认 en；其他模板保持 larkConfig.sourceLang
    const sourceLang = templateKey === 'comprehensive'
      ? (larkConfig.comprehensiveSourceLang || 'en')
      : larkConfig.sourceLang;
    const targetLangs = bgFiles.map(f => path.parse(f).name).filter(l => l !== sourceLang);

    let translationsMap = { [sourceLang]: sourceData };

    // 优先使用表格里已有的翻译，缺失的语种才调翻译 API
    const existing = sourceData.existingTranslations ?? {};
    const missingLangs = targetLangs.filter(l => {
      const e = existing[l.toLowerCase()];
      return !e || !e.mainTitle;
    });

    for (const targetLang of targetLangs) {
      const data = existing[targetLang.toLowerCase()];
      if (data) translationsMap[targetLang] = data;
    }

    if (missingLangs.length > 0) {
      console.log(`\n正在翻译 ${missingLangs.length} 种语言（${missingLangs.join(', ')}）...`);
      const translated = await translateComprehensiveData(sourceData, missingLangs, sourceLang);
      translationsMap = { ...translationsMap, ...translated };

      console.log('\n回填翻译结果到 Lark 表格...');
      const writtenCount = await writeBackTranslationsToLark(
        sourceData, translationsMap,
        accessToken, larkConfig.spreadsheetToken, sheetId, sourceLang,
        { includeSourceLang: true }
      );
      console.log(`  ✅ 已回填 ${writtenCount} 种语言（第 3 行起，A 列为语言代码）`);
    } else {
      console.log(`\n✅ 已使用表格中的现有翻译（${Object.keys(translationsMap).length} 种语言）`);
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${bgFiles.length} 个语种）：`);
    for (const bgFile of bgFiles) {
      const lang = path.parse(bgFile).name;
      const bgPath = path.join(bgDir, bgFile);
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.jpg`);
      const posterPayload = buildComprehensivePosterPayload(sourceData, translationsMap, lang, copyConfig, templateKey);
      await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, bgFiles);
    return;
  }

  // ── Classic NBA 流程 ──────────────────────────────────────
  console.log('\n从 Lark NBA 表格拉取主文案和比赛链接...');
  const larkConfig = loadLarkConfig();
  const accessToken = await getLarkTenantAccessToken(larkConfig);
  const classicData = await fetchClassicDataFromLark(larkConfig, accessToken, larkConfig.sourceLang);
  const teamsCsvPath = templateConfig.teamsCsv || TEAMS_CSV;
  const teamsMap = loadTeams(teamsCsvPath);
  let gamesRows = await Promise.all(
    classicData.matchInputs.map(item => resolveClassicMatchFromPolymarketInput(item, teamsMap))
  );
  gamesRows = normalizeGameRowsTeamIds(gamesRows, teamsMap);
  console.log(`  ✅ 已解析 ${gamesRows.length} 场比赛链接`);

  const matchWriteCount = await writeBackClassicMatchesToLark(
    gamesRows,
    classicData.headerIndexes,
    classicData.larkContext,
    classicData.sourceRowNumber
  );
  if (matchWriteCount > 0) {
    console.log(`  ✅ 已回填 Lark 表格 ${matchWriteCount} 场比赛的主客队、日期和赔率`);
  }

  const bgDir = templateConfig.bgDir || BG_DIR;
  const NBA_ALL_TARGET_LANGS = ['zh-TW', 'en', 'ja', 'de', 'es', 'fr', 'pt', 'vi'];

  // 竖版所有语言共用一张背景图（取 bgDir 下第一张图）
  const nbaBgFileNames = scanBgFiles(bgDir);
  const verticalBgPath = path.join(bgDir, nbaBgFileNames[0]);
  const sourceLang = String(larkConfig.sourceLang || 'zh-CN').trim();
  const targetLangs = NBA_ALL_TARGET_LANGS.filter(l => l !== sourceLang);
  const langsToGenerate = [sourceLang, ...targetLangs];

  let translationsMap = { [sourceLang]: classicData.sourceData };
  if (targetLangs.length > 0) {
    console.log('\n翻译 NBA 标题、副标题和 footer...');
    const translated = await translateClassicTitles(classicData.sourceData, targetLangs, sourceLang);
    translationsMap = { ...translationsMap, ...translated };

    console.log('\n回填翻译结果到 Lark 表格...');
    const translationWrittenCount = await writeBackClassicTranslationsToLark(
      classicData.sourceData,
      translationsMap,
      accessToken,
      classicData.larkContext.spreadsheetToken,
      classicData.larkContext.sheetId,
      sourceLang,
      classicData.headerIndexes
    );
    if (translationWrittenCount > 0) {
      console.log(`  ✅ 已回填 Lark 表格 ${translationWrittenCount} 行翻译`);
    }
  }

  validateWinRates(gamesRows);
  const verticalTemplate = fs.readFileSync(templateConfig.file, 'utf8');
  const horizontalFile = templateConfig.horizontalFile;
  const horizontalTemplate = horizontalFile && fs.existsSync(horizontalFile)
    ? fs.readFileSync(horizontalFile, 'utf8')
    : '';

  // 横版背景：所有语言共用一张图（取 horizontalBgDir 下第一张图）
  let horizontalBgPath = '';
  const horizontalBgDir = templateConfig.horizontalBgDir;
  if (horizontalTemplate && horizontalBgDir && fs.existsSync(horizontalBgDir)) {
    const horizontalBgFiles = fs.readdirSync(horizontalBgDir).filter(f => /\.(png|jpg|jpeg)$/i.test(f));
    if (horizontalBgFiles.length > 0) {
      horizontalBgPath = path.join(horizontalBgDir, horizontalBgFiles[0]);
    }
  }
  if (horizontalTemplate && !horizontalBgPath) {
    console.warn('⚠️  未找到横版背景图（backgrounds-NBA-horizontal/），横版将使用纯黑背景');
  }

  const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
  const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

  const browser = await launchBrowser();
  const page = await browser.newPage();

  const generatedFiles = [];

  console.log(`\n生成竖版 1:1 海报（${langsToGenerate.length} 个语种）：`);
  for (const lang of langsToGenerate) {
    const bgPath = verticalBgPath;
    const fileName = `${outputPrefixWithDate}_1-1_${lang}.jpg`;
    const outputPath = path.join(dateDir, fileName);
    const posterPayload = buildClassicPosterPayload(gamesRows, teamsMap, lang, classicData.sourceData, translationsMap, copyConfig, { layout: 'vertical' });
    await generatePoster(page, verticalTemplate, posterPayload, bgPath, outputPath, {
      viewportWidth: 1200,
      viewportHeight: 1200,
      outputWidth: 1200,
      outputHeight: 1200
    });
    generatedFiles.push(fileName);
  }

  if (horizontalTemplate) {
    console.log(`\n生成横版 2:1 海报（${langsToGenerate.length} 个语种）：`);
    for (const lang of langsToGenerate) {
      const fileName = `${outputPrefixWithDate}_2-1_${lang}.jpg`;
      const outputPath = path.join(dateDir, fileName);
      const posterPayload = buildClassicPosterPayload(gamesRows, teamsMap, lang, classicData.sourceData, translationsMap, copyConfig, { layout: 'horizontal' });
      await generatePoster(page, horizontalTemplate, posterPayload, horizontalBgPath, outputPath, {
        viewportWidth: 2400,
        viewportHeight: 1200,
        outputWidth: 2400,
        outputHeight: 1200
      });
      generatedFiles.push(fileName);
    }
  }

  await browser.close();

  // 打包所有生成的文件到同一个 zip
  const zipName = `${outputPrefixWithDate}.zip`;
  const zipPath = path.join(dateDir, zipName);
  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
  const quotedFiles = generatedFiles.map(f => `"${f}"`).join(' ');
  execSync(`cd "${dateDir}" && zip "${zipName}" ${quotedFiles}`);
  const zipKB = Math.round(fs.statSync(zipPath).size / 1024);
  console.log(`\n📦 ${zipName} (${zipKB}KB)`);
  console.log(`所有图片已保存到：${dateDir}\n`);
}

main().catch(err => {
  console.error('❌ 生成失败：', err.message);
  process.exit(1);
});
