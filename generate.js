const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const { execSync } = require('child_process');

const MAX_SIZE_KB = 300;
const POLYMARKET_BASE_URL = 'https://gamma-api.polymarket.com';

// ── 压缩图片到目标大小以内（JPEG 二分查找质量参数）──
async function compressToTarget(inputPath, outputPath, maxKB = MAX_SIZE_KB) {
  let lo = 20, hi = 90, bestQuality = 60;
  // 先快速测试最高质量，如果已经够小就直接用
  for (let i = 0; i < 6; i++) {
    const quality = Math.round((lo + hi) / 2);
    const buf = await sharp(inputPath).jpeg({ quality }).toBuffer();
    const sizeKB = buf.length / 1024;
    if (sizeKB <= maxKB) {
      bestQuality = quality;
      lo = quality + 1; // 尝试更高质量
    } else {
      hi = quality - 1; // 降低质量
    }
  }
  await sharp(inputPath).jpeg({ quality: bestQuality }).toFile(outputPath);
  const finalKB = Math.round(fs.statSync(outputPath).size / 1024);
  return { quality: bestQuality, sizeKB: finalKB };
}

const BASE_DIR = __dirname;
const TEAMS_CSV    = path.join(BASE_DIR, 'teams.csv');
const FOOTBALL_TEAMS_CSV = path.join(BASE_DIR, 'football_teams.csv');
const F1_TEAMS_CSV = path.join(BASE_DIR, 'teams_f1.csv');
const F1_ICON_DIR  = path.join(BASE_DIR, 'assets', 'logos', 'F1 车队');
const BG_DIR       = path.join(BASE_DIR, 'backgrounds-NBA');
const WORLD_CUP_BG_DIR = path.join(BASE_DIR, 'backgrounds- football');
const OUTPUT_DIR  = path.join(BASE_DIR, 'output');
const POSTER_COPY_CONFIG = path.join(BASE_DIR, 'poster.copy.json');

const LARK_CONFIG = path.join(BASE_DIR, 'lark.config.json');
const TEMPLATE_CONFIGS = {
  classic: {
    aliases: ['classic', 'default', '标准', '默认'],
    file: path.join(BASE_DIR, 'poster.html'),
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
  f1: {
    aliases: ['f1', 'f1赛车', 'F1', 'formula1', 'formula-1'],
    file: path.join(BASE_DIR, 'poster.f1.html'),
    outputPrefix: 'F1',
    outputSubDir: 'F1 车队'
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
        '  f1             F1车队海报模版'
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
  const logoPath = path.join(BASE_DIR, 'NBA_icon', `${logoName}.png`);
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

function normalizeOutcomeToken(value) {
  return String(value ?? '')
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
    ['usa', ['unitedstates', 'unitedstatesofamerica', 'us']],
    ['south_korea', ['korearepublic', 'republicofkorea', 'korea']],
    ['england', ['englandnationalteam']],
    ['netherlands', ['holland']],
    ['ivory_coast', ['cotedivoire']]
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
      const displayText = String(team?.['zh-CN'] ?? outcome).trim();

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
    const displayText = String(team?.['zh-CN'] ?? teamLabel).trim();

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
  return fields.some(item => String(item ?? '').includes(dateYmd));
}

function getTeamEnglishName(teamId, teamsMap) {
  const en = String(teamsMap[teamId]?.en ?? '').trim();
  if (!en) {
    throw new Error(`球队 ${teamId} 缺少英文名，无法自动匹配 Polymarket`);
  }
  return en;
}

function scoreMarketCandidate(market, dateYmd, homeAliases, awayAliases) {
  const question = String(market?.question ?? '');
  const slug = String(market?.slug ?? '');
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

  let score = 0;
  if (nbaHint) score += 40;
  if (homeMatched) score += 40;
  if (awayMatched) score += 40;
  if (dateMatched) score += 40;
  if (outcomeHome && outcomeAway) score += 40;
  if (textHome && textAway) score += 20;
  if (market?.active === true) score += 5;
  score += Math.min(Number(market?.volumeNum ?? market?.volume ?? 0) / 10000, 10);
  return {
    score: homeMatched && awayMatched && dateMatched ? score : -1,
    previewScore: score,
    reason: homeMatched && awayMatched && dateMatched ? 'ok' : 'team/date mismatch',
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

async function resolvePolymarketMarketByGame(row, teamsMap, activeNbaMarkets, searchCache) {
  if (row.polymarket_slug) {
    return { slug: row.polymarket_slug, market: null };
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
    const result = scoreMarketCandidate(market, dateYmd, homeAliases, awayAliases);
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
      `请检查表格中的日期和球队 ID，或手动确认这场比赛在 Polymarket 上是否存在。` +
      candidateHint
    );
  }

  const slug = String(best.slug).trim();
  const resolved = { slug, market: best };
  searchCache.set(cacheKey, resolved);
  return resolved;
}

async function enrichRowsWithPolymarketOdds(gamesRows, teamsMap) {
  const cache = new Map();
  const searchCache = new Map();
  const activeNbaMarkets = await fetchPolymarketActiveNbaMarkets();
  const nextRows = [];
  let enrichedCount = 0;
  let autoMatchedCount = 0;

  for (const row of gamesRows) {
    const inputSlug = String(row.polymarket_slug ?? '').trim();
    const { slug, market: resolvedMarket } = await resolvePolymarketMarketByGame(
      row,
      teamsMap,
      activeNbaMarkets,
      searchCache
    );
    if (!inputSlug && slug) autoMatchedCount++;

    let market = resolvedMarket ?? cache.get(slug);
    if (!market) {
      market = await fetchPolymarketMarketBySlug(slug);
      cache.set(slug, market);
    }

    const outcomes = parseJsonArrayField(market.outcomes, 'outcomes', slug).map(String);
    const outcomePrices = parseJsonArrayField(market.outcomePrices, 'outcomePrices', slug);
    if (outcomes.length === 0 || outcomes.length !== outcomePrices.length) {
      throw new Error(`Polymarket 数据异常（${slug}）：outcomes 与 outcomePrices 数量不一致`);
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
    const { homeWin, awayWin } = roundWinRatesToIntegers(rawHomeWin, rawAwayWin);
    if (!Number.isFinite(homeWin) || !Number.isFinite(awayWin)) {
      throw new Error(`Polymarket 概率解析失败（${slug}）：outcomePrices=${JSON.stringify(outcomePrices)}`);
    }

    nextRows.push({
      ...row,
      polymarket_slug: slug,
      home_win: String(homeWin),
      away_win: String(awayWin),
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
  const range = encodeURIComponent(`${resolvedSheetId}!A1:L35`);
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
  const row = values[1].map(cell => String(cell ?? '').trim());

  function col(name) {
    const idx = headers.indexOf(name);
    return idx !== -1 ? row[idx] : '';
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
      return {
        text: col(String(index)),
        percent: Number.isFinite(percentRaw) ? percentRaw : 50
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

  return {
    mainTitle: col('main title'),
    subTitle: col('sub title'),
    footer: col('footer'),
    cards
  };
}

// ── 综合事件：调用 MyMemory 免费翻译（无需 API Key）──────────
async function translateOneText(text, fromLang, toLang) {
  if (!text) return text;
  const params = new URLSearchParams({ q: text, langpair: `${fromLang}|${toLang}` });
  const res = await fetch(`https://api.mymemory.translated.net/get?${params}`);
  if (!res.ok) {
    throw new Error(`MyMemory 翻译请求失败（${fromLang}→${toLang}）：HTTP ${res.status}`);
  }
  const data = await res.json();
  if (data.responseStatus !== 200) {
    throw new Error(`MyMemory 翻译失败（${fromLang}→${toLang}）：${data.responseDetails}`);
  }
  return String(data.responseData?.translatedText ?? text);
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
        text: translated[3 + i],
        percent: card.percent
      }))
    };
    console.log(' ✅');
  }

  return result;
}

// ── 综合事件：翻译结果回填到 Lark 表格（第 3 行起，仅回填非源语言）──
async function writeBackTranslationsToLark(
  sourceData,
  translationsMap,
  accessToken,
  spreadsheetToken,
  sheetId,
  sourceLang = 'zh-CN'
) {
  const sourceLangNormalized = String(sourceLang ?? '').trim().toLowerCase();
  const langs = Object.keys(translationsMap)
    .filter(lang => String(lang ?? '').trim().toLowerCase() !== sourceLangNormalized);

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

// ── 综合事件：构建海报 payload ──────────────────────────────
function buildComprehensivePosterPayload(sourceData, translationsMap, lang, copyConfig, templateKey = 'comprehensive') {
  const templateCopy = copyConfig?.[templateKey] ?? copyConfig?.comprehensive ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };

  // Use translation for this lang; fall back to source (zh-CN) data
  const translated = translationsMap[lang] ?? sourceData;

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
      cardTextFontSize: Number(mergedCopy.cardTextFontSize ?? 40),
      cardTextLineHeight: Number(mergedCopy.cardTextLineHeight ?? 1.2),
      cards: (translated.cards ?? sourceData.cards).map((card, i) => ({
        text: String(card.text ?? '').trim(),
        image: templateKey === 'worldcup'
          ? (String(sourceData.cards?.[i]?.image ?? '').trim()
            || resolveWorldCupLogoPath(sourceData.cards?.[i]?.text ?? card.text, i))
          : String(defaultCopy.cards?.[i]?.image ?? '').trim(),
        valueLabel: String(mergedCopy.outcomeLabel ?? 'Yes')
      }))
    }
  };
}

// ── 生成单张海报 ──────────────────────────────────────────
async function generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath) {
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

  const tmpPng = outputPath.replace(/\.png$/i, '') + '_raw.png';
  await page.screenshot({
    path: tmpPng,
    clip: { x: 0, y: 0, width: 1200, height: 1200 }
  });

  fs.unlinkSync(tmpFile);

  const jpgPath = outputPath.replace(/\.png$/i, '.jpg');
  const { quality, sizeKB } = await compressToTarget(tmpPng, jpgPath);
  fs.unlinkSync(tmpPng);
  console.log(`  ✅ ${path.basename(jpgPath)} (${sizeKB}KB, quality=${quality})`);
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
  return puppeteer.launch({
    headless: 'new',
    executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--allow-file-access-from-files']
  });
}

function buildZip(dateDir, outputPrefixWithDate, bgFiles) {
  const zipName = `${outputPrefixWithDate}.zip`;
  const zipPath = path.join(dateDir, zipName);
  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
  const files = bgFiles.map(f => `${outputPrefixWithDate}_${path.parse(f).name}.jpg`).join(' ');
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
          console.log(`     ${polyTeams.map(t => `${t._rawName ?? t.teamId}(${t.percent}%)`).join(', ')}`);
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
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.png`);
      const posterPayload = buildF1PosterPayload(sourceData, translationsMap, f1TeamsMap, lang, copyConfig);
      await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, langsToGenerate.map(l => `${l}.png`));
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
    const sourceLang = larkConfig.sourceLang;
    const targetLangs = bgFiles.map(f => path.parse(f).name).filter(l => l !== sourceLang);

    let translationsMap = { [sourceLang]: sourceData };

    if (targetLangs.length > 0) {
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')}）...`);
      const translated = await translateComprehensiveData(sourceData, targetLangs, larkConfig.sourceLang);
      translationsMap = { ...translationsMap, ...translated };

      console.log('\n回填翻译结果到 Lark 表格...');
      const writtenCount = await writeBackTranslationsToLark(
        sourceData, translationsMap,
        accessToken, larkConfig.spreadsheetToken, sheetId, larkConfig.sourceLang
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

    console.log(`\n生成海报（${bgFiles.length} 个语种）：`);
    for (const bgFile of bgFiles) {
      const lang = path.parse(bgFile).name;
      const bgPath = path.join(bgDir, bgFile);
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.png`);
      const posterPayload = buildComprehensivePosterPayload(sourceData, translationsMap, lang, copyConfig, templateKey);
      await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
    }

    await browser.close();
    buildZip(dateDir, outputPrefixWithDate, bgFiles);
    return;
  }

  // ── Classic NBA 流程 ──────────────────────────────────────
  console.log('\n从 Lark 普通电子表格拉取比赛数据...');
  const sheetData = await fetchGamesFromLarkSheets();
  let gamesRows = sheetData.rows;
  console.log(`  ✅ 共 ${gamesRows.length} 场比赛`);

  const teamsMap  = loadTeams(TEAMS_CSV);
  const { rows: enrichedRows, enrichedCount, autoMatchedCount } = await enrichRowsWithPolymarketOdds(gamesRows, teamsMap);
  gamesRows = enrichedRows;
  if (enrichedCount > 0) {
    console.log(`  ✅ 已从 Polymarket 自动更新 ${enrichedCount} 场赔率`);
  }
  if (autoMatchedCount > 0) {
    console.log(`  ✅ 其中 ${autoMatchedCount} 场由主客队+日期自动匹配到 Polymarket 市场`);
  }

  const writtenCount = await writeBackWinRatesToLark(gamesRows, sheetData.headerIndexes, sheetData.larkContext);
  if (writtenCount > 0) {
    console.log(`  ✅ 已回填 Lark 表格 ${writtenCount} 场赔率`);
  }

  validateWinRates(gamesRows);
  const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');

  const bgFiles = scanBgFiles();
  const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const dateDir = prepareOutputDir(date, templateConfig.outputSubDir);
  const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

  const browser = await launchBrowser();
  const page = await browser.newPage();
  await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

  console.log(`\n生成海报（${bgFiles.length} 个语种）：`);
  for (const bgFile of bgFiles) {
    const lang = path.parse(bgFile).name;
    const bgPath = path.join(BG_DIR, bgFile);
    const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.png`);
    const posterPayload = buildPosterPayload(gamesRows, teamsMap, lang, templateKey, copyConfig);
    await generatePoster(page, htmlTemplate, posterPayload, bgPath, outputPath);
  }

  await browser.close();
  buildZip(dateDir, outputPrefixWithDate, bgFiles);
}

main().catch(err => {
  console.error('❌ 生成失败：', err.message);
  process.exit(1);
});
