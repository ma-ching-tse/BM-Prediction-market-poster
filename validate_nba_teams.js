/**
 * 验证 teams.csv 里所有 NBA 球队的 Polymarket 名称匹配情况
 * 测试两条解析路径：
 *   1. resolveTeamId      — outcomes 里的写法（精确/尾词匹配）
 *   2. resolveTeamIdFromTextSegment — 问题文本里的写法（includes 模糊匹配）
 */

const fs = require('fs');
const path = require('path');

// ── 从 generate.js 复制的核心函数 ──

function normalizeOutcomeToken(value) {
  return String(value ?? '')
    .normalize('NFKD')
    .replace(/\p{M}+/gu, '')
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, '');
}

function loadTeams(csvPath) {
  const content = fs.readFileSync(csvPath, 'utf8');
  const lines = content.split('\n').filter(Boolean);
  const headers = lines[0].split(',');
  const map = {};
  for (const line of lines.slice(1)) {
    const cols = line.split(',');
    const obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = (cols[i] ?? '').trim(); });
    if (obj.id) map[obj.id] = obj;
  }
  return map;
}

// 精确/尾词匹配（对应 outcomes 路径，含输入尾词反向匹配）
function resolveTeamId(input, teamsMap) {
  const raw = String(input ?? '').trim();
  if (!raw) return null;
  if (teamsMap[raw]) return raw;
  const lowered = raw.toLowerCase();
  if (teamsMap[lowered]) return lowered;
  const normalizedRaw = normalizeOutcomeToken(raw);
  if (!normalizedRaw) return null;
  for (const [teamId, team] of Object.entries(teamsMap)) {
    const candidates = [teamId, team.en, team['zh-CN'], team['zh-TW'], team.ja];
    for (const candidate of candidates) {
      if (!candidate) continue;
      const nc = normalizeOutcomeToken(candidate);
      if (nc && nc === normalizedRaw) return teamId;
      const parts = String(candidate).trim().split(/\s+/);
      if (parts.length > 1) {
        const tail = normalizeOutcomeToken(parts[parts.length - 1]);
        if (tail && tail === normalizedRaw) return teamId;
      }
    }
  }
  // 输入本身是多词时，用输入的尾词反向匹配
  const inputParts = raw.trim().split(/\s+/);
  if (inputParts.length > 1) {
    const inputTail = normalizeOutcomeToken(inputParts[inputParts.length - 1]);
    if (inputTail) {
      if (teamsMap[inputTail]) return inputTail;
      for (const [teamId, team] of Object.entries(teamsMap)) {
        if (normalizeOutcomeToken(team.en) === inputTail) return teamId;
      }
    }
  }
  return null;
}

// includes 模糊匹配（对应问题文本路径）
function resolveTeamIdFromTextSegment(text, teamsMap) {
  const normalizedText = normalizeOutcomeToken(text);
  if (!normalizedText) return null;
  let best = { teamId: null, score: 0 };
  for (const [teamId, team] of Object.entries(teamsMap)) {
    const candidates = [teamId, team.en, team['zh-CN'], team['zh-TW'], team.ja];
    for (const candidate of candidates) {
      if (!candidate) continue;
      const nc = normalizeOutcomeToken(candidate);
      if (!nc) continue;
      if (normalizedText.includes(nc) || nc.includes(normalizedText)) {
        if (nc.length > best.score) best = { teamId, score: nc.length };
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

// ── Polymarket 常用的 NBA 完整写法（城市 + 队名）──
const NBA_FULL_NAMES = {
  hawks:        'Atlanta Hawks',
  celtics:      'Boston Celtics',
  nets:         'Brooklyn Nets',
  hornets:      'Charlotte Hornets',
  bulls:        'Chicago Bulls',
  cavaliers:    'Cleveland Cavaliers',
  mavericks:    'Dallas Mavericks',
  nuggets:      'Denver Nuggets',
  pistons:      'Detroit Pistons',
  warriors:     'Golden State Warriors',
  rockets:      'Houston Rockets',
  pacers:       'Indiana Pacers',
  clippers:     'Los Angeles Clippers',
  lakers:       'Los Angeles Lakers',
  grizzlies:    'Memphis Grizzlies',
  heat:         'Miami Heat',
  bucks:        'Milwaukee Bucks',
  timberwolves: 'Minnesota Timberwolves',
  pelicans:     'New Orleans Pelicans',
  knicks:       'New York Knicks',
  thunder:      'Oklahoma City Thunder',
  magic:        'Orlando Magic',
  '76ers':      'Philadelphia 76ers',
  suns:         'Phoenix Suns',
  blazers:      'Portland Trail Blazers',
  kings:        'Sacramento Kings',
  spurs:        'San Antonio Spurs',
  raptors:      'Toronto Raptors',
  jazz:         'Utah Jazz',
  wizards:      'Washington Wizards',
};

// ── 主逻辑 ──

const CSV = path.join(__dirname, 'teams.csv');
const teamsMap = loadTeams(CSV);
const teams = Object.values(teamsMap);

console.log(`\n📋 共 ${teams.length} 支 NBA 球队，开始验证...\n`);

const exactFails  = [];  // outcomes 路径（精确匹配）失败
const fuzzyFails  = [];  // 问题文本路径（模糊匹配）失败

for (const team of teams) {
  const id = team.id;
  const fullName = NBA_FULL_NAMES[id];

  // 测试 1：短昵称（outcomes 最常见写法）
  const shortMatch = resolveTeamId(team.en, teamsMap);
  const shortOk = shortMatch === id;

  // 测试 2：完整城市+队名 via resolveTeamId（outcomes 精确路径）
  const exactFullMatch = fullName ? resolveTeamId(fullName, teamsMap) : id;
  const exactFullOk = !fullName || exactFullMatch === id;

  // 测试 3：完整城市+队名 via resolveTeamIdFromTextSegment（问题文本路径）
  const fuzzyFullMatch = fullName ? resolveTeamIdFromTextSegment(fullName, teamsMap) : id;
  const fuzzyFullOk = !fullName || fuzzyFullMatch === id;

  if (!shortOk || !exactFullOk) {
    exactFails.push({ team, shortMatch, exactFullMatch, fullName });
  }
  if (!fuzzyFullOk) {
    fuzzyFails.push({ team, fuzzyFullMatch, fullName });
  }
}

// ── 输出结果 ──

if (exactFails.length === 0) {
  console.log('✅ [outcomes 路径] 全部 30 支球队匹配正常（短昵称 + 精确匹配均通过）\n');
} else {
  console.log(`⚠️  [outcomes 路径] ${exactFails.length} 支球队存在问题：\n`);
  for (const { team, shortMatch, exactFullMatch, fullName } of exactFails) {
    console.log(`  ❌ [${team.id}] (en: "${team.en}", full: "${fullName}")`);
    if (!shortMatch) console.log(`       短昵称 "${team.en}" → 无法匹配`);
    if (fullName && exactFullMatch !== team.id) console.log(`       完整名 "${fullName}" → 匹配到 "${exactFullMatch ?? '无'}"（应为 ${team.id}）`);
  }
  console.log('');
}

if (fuzzyFails.length === 0) {
  console.log('✅ [问题文本路径] 全部 30 支球队匹配正常（城市+队名 includes 匹配均通过）\n');
} else {
  console.log(`⚠️  [问题文本路径] ${fuzzyFails.length} 支球队存在问题：\n`);
  for (const { team, fuzzyFullMatch, fullName } of fuzzyFails) {
    console.log(`  ❌ [${team.id}] "${fullName}" → 匹配到 "${fuzzyFullMatch ?? '无'}"（应为 ${team.id}）`);
  }
  console.log('');
}

// 特殊边缘案例
console.log('🔍 边缘案例验证：');
const edgeCases = [
  ['76ers', 'Philadelphia 76ers'],
  ['76ers', '76ers'],
  ['clippers', 'LA Clippers'],
  ['lakers', 'LA Lakers'],
  ['blazers', 'Trail Blazers'],
  ['timberwolves', "Minnesota T-Wolves"],
];
for (const [expectedId, input] of edgeCases) {
  const r1 = resolveTeamId(input, teamsMap);
  const r2 = resolveTeamIdFromTextSegment(input, teamsMap);
  const exact = r1 === expectedId ? '✅' : `❌(精确→${r1 ?? '无'})`;
  const fuzzy = r2 === expectedId ? '✅' : `❌(模糊→${r2 ?? '无'})`;
  console.log(`  "${input}" → 精确:${exact}  模糊:${fuzzy}`);
}
