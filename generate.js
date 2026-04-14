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
const TEAMS_CSV   = path.join(BASE_DIR, 'teams.csv');
const BG_DIR      = path.join(BASE_DIR, 'backgrounds');
const OUTPUT_DIR  = path.join(BASE_DIR, 'output');
const POSTER_COPY_CONFIG = path.join(BASE_DIR, 'poster.copy.json');

const LARK_CONFIG = path.join(BASE_DIR, 'lark.config.json');
const TEMPLATE_CONFIGS = {
  classic: {
    aliases: ['classic', 'default', '标准', '默认'],
    file: path.join(BASE_DIR, 'poster.html'),
    outputPrefix: 'NBA'
  },
  comprehensive: {
    aliases: ['comprehensive', 'event', '综合事件', '综合事件模版'],
    file: path.join(BASE_DIR, 'poster.comprehensive-event.html'),
    outputPrefix: 'NBA_综合事件'
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
        '  comprehensive  综合事件模版'
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
async function fetchComprehensiveEventsFromLark(config, accessToken) {
  const sheetId = String(config.comprehensiveSheetId ?? '').trim();
  if (!sheetId) {
    throw new Error('lark.config.json 缺少 comprehensiveSheetId');
  }

  const spreadsheetToken = config.spreadsheetToken;
  const range = encodeURIComponent(`${sheetId}!A1:I2`);
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

  return {
    mainTitle: col('main title'),
    subTitle: col('sub title'),
    footer: col('footer'),
    cards: [
      { text: col('1'), percent: Number(col('percent_1')) || 50 },
      { text: col('2'), percent: Number(col('percent_2')) || 50 },
      { text: col('3'), percent: Number(col('percent_3')) || 50 }
    ]
  };
}

// ── 综合事件：调用 Claude API 翻译 ──────────────────────────
async function translateComprehensiveData(sourceData, targetLangs, apiKey) {
  const langLabels = {
    'zh-TW': '繁體中文（台灣）',
    'en': 'English',
    'ja': '日本語',
    'es': 'Español',
    'pt': 'Português',
    'de': 'Deutsch',
    'fr': 'Français',
    'vi': 'Tiếng Việt'
  };

  const targets = targetLangs.map(l => `${l}（${langLabels[l] || l}）`).join('、');

  const prompt = `你是一名专业的本地化翻译员，专注于加密货币交易平台的营销文案翻译。

请将以下简体中文内容翻译成：${targets}

原始内容：
${JSON.stringify(sourceData, null, 2)}

翻译要求：
- mainTitle：海报主标题，语气要吸引人，简洁有力
- subTitle：副标题，自然流畅
- footer：产品引导文案，简短清晰
- cards[].text：预测市场问题，保持问句形式
- cards[].percent：直接复制原值，不翻译
- 不修改 JSON 结构和字段名
- 直接输出 JSON，不要加任何解释

输出格式（只输出这个 JSON 对象）：
{
  "zh-TW": { "mainTitle": "...", "subTitle": "...", "footer": "...", "cards": [{"text": "...", "percent": 72}, ...] },
  "en": { ... }
}`;

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 4096,
      messages: [{ role: 'user', content: prompt }]
    })
  });

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Claude 翻译 API 失败：HTTP ${res.status}，${text.slice(0, 200)}`);
  }

  const result = await res.json();
  const raw = result.content?.[0]?.text ?? '';

  const jsonMatch = raw.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('Claude 翻译返回格式异常：未找到 JSON 块');
  }

  try {
    return JSON.parse(jsonMatch[0]);
  } catch (e) {
    throw new Error(`Claude 翻译返回的 JSON 解析失败：${e.message}\n内容片段：${raw.slice(0, 300)}`);
  }
}

// ── 综合事件：构建海报 payload ──────────────────────────────
function buildComprehensivePosterPayload(sourceData, translationsMap, lang, copyConfig) {
  const templateCopy = copyConfig?.comprehensive ?? {};
  const defaultCopy = templateCopy?.default ?? {};
  const langCopy = templateCopy?.[lang] ?? {};
  const mergedCopy = { ...defaultCopy, ...langCopy };

  // Use translation for this lang; fall back to source (zh-CN) data
  const translated = translationsMap[lang] ?? sourceData;

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
      cards: (translated.cards ?? sourceData.cards).map(card => ({
        text: String(card.text ?? '').trim(),
        image: '',
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
function scanBgFiles() {
  const bgFiles = fs.readdirSync(BG_DIR).filter(f => f.endsWith('.png') || f.endsWith('.jpg'));
  if (bgFiles.length === 0) {
    console.error('❌ backgrounds/ 目录下没有找到背景图，请先添加语种背景图（如 zh-CN.png）');
    process.exit(1);
  }
  return bgFiles;
}

function prepareOutputDir(date) {
  const dateDir = path.join(OUTPUT_DIR, date);
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);
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

  // ── 综合事件流程 ──────────────────────────────────────────
  if (templateKey === 'comprehensive') {
    console.log('\n从 Lark 综合事件表格拉取数据...');
    const larkConfig = loadLarkConfig();
    const accessToken = await getLarkTenantAccessToken(larkConfig);
    const sourceData = await fetchComprehensiveEventsFromLark(larkConfig, accessToken);
    console.log('  ✅ 已读取综合事件数据');

    const bgFiles = scanBgFiles();
    const sourceLang = larkConfig.sourceLang;
    const targetLangs = bgFiles.map(f => path.parse(f).name).filter(l => l !== sourceLang);

    let translationsMap = { [sourceLang]: sourceData };

    if (targetLangs.length > 0) {
      const apiKey = larkConfig.anthropicApiKey || String(process.env.ANTHROPIC_API_KEY ?? '').trim();
      if (!apiKey) {
        throw new Error(
          '需要 Anthropic API Key 进行多语言翻译，' +
          '请在 lark.config.json 的 anthropicApiKey 字段或环境变量 ANTHROPIC_API_KEY 中配置'
        );
      }
      console.log(`\n正在翻译 ${targetLangs.length} 种语言（${targetLangs.join(', ')}）...`);
      const translated = await translateComprehensiveData(sourceData, targetLangs, apiKey);
      translationsMap = { ...translationsMap, ...translated };
      console.log('  ✅ 翻译完成');
    }

    const htmlTemplate = fs.readFileSync(templateConfig.file, 'utf8');
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const dateDir = prepareOutputDir(date);
    const outputPrefixWithDate = `${templateConfig.outputPrefix}_${date}`;

    const browser = await launchBrowser();
    const page = await browser.newPage();
    await page.setViewport({ width: 1200, height: 1200, deviceScaleFactor: 2 });

    console.log(`\n生成海报（${bgFiles.length} 个语种）：`);
    for (const bgFile of bgFiles) {
      const lang = path.parse(bgFile).name;
      const bgPath = path.join(BG_DIR, bgFile);
      const outputPath = path.join(dateDir, `${outputPrefixWithDate}_${lang}.png`);
      const posterPayload = buildComprehensivePosterPayload(sourceData, translationsMap, lang, copyConfig);
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
  const dateDir = prepareOutputDir(date);
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
