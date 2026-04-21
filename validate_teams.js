/**
 * 验证 football_teams.csv 里所有球队的 Polymarket 名称匹配情况
 * 测试常见的 Polymarket 写法变体，找出会匹配失败的球队
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
];

function resolve(outcomeText, teamsMap) {
  const n = normalizeOutcomeToken(outcomeText);
  if (!n) return null;
  for (const team of Object.values(teamsMap)) {
    const aliases = [team.id, team.en, team['zh-CN'], team['zh-TW'], team.ja]
      .map(v => normalizeOutcomeToken(v)).filter(Boolean);
    if (aliases.some(a => a === n)) return team;
  }
  for (const [teamId, aliases] of manualAliases) {
    if (aliases.includes(n) && teamsMap[teamId]) return teamsMap[teamId];
  }
  return null;
}

// ── 常见 Polymarket 名称变体规则 ──
// 生成每支球队可能在 Polymarket 出现的名称变体

function generatePolymarketVariants(team) {
  const en = team.en || '';
  const id = team.id || '';
  const variants = new Set();

  // 原始 en 和 id
  variants.add(en);

  // 去掉常见前后缀
  const prefixes = ['FC ', 'CF ', 'SC ', 'AC ', 'AS ', 'RC ', 'RCD ', 'SS ', 'SSC ', 'SV ', 'VfB ', 'VfL ', 'TSG ', 'CA ', 'US ', 'AJ ', 'UD ', 'OGC ', 'LOSC ', 'Stade ', 'Olympique '];
  const suffixes = [' FC', ' CF', ' SC', ' AC', ' CFC', ' AFC', ' SV', ' 1907', ' 05', ' 04', ' 07', ' Calcio', ' UD', ' SCO', ' AC'];

  let stripped = en;
  for (const p of prefixes) {
    if (stripped.startsWith(p)) { variants.add(stripped.slice(p.length)); stripped = stripped.slice(p.length); break; }
  }
  for (const s of suffixes) {
    if (en.endsWith(s)) { variants.add(en.slice(0, -s.length)); break; }
  }

  // id 转人类可读（下划线→空格→首字母大写）
  const idReadable = id.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
  variants.add(idReadable);

  // 特殊已知变体
  const knownVariants = {
    psg: ['Paris Saint-Germain FC', 'PSG'],
    atletico_madrid: ['Club Atlético de Madrid', 'Atlético de Madrid'],
    athletic_bilbao: ['Athletic Club', 'Athletic Club Bilbao'],
    celta_vigo: ['Celta de Vigo', 'RC Celta'],
    monchengladbach: ["Borussia Mönchengladbach", "Borussia Monchengladbach"],
    real_madrid: ['Real Madrid'],
    barcelona: ['Barcelona'],
    inter_milan: ['Inter', 'FC Internazionale', 'Internazionale'],
    sporting_cp: ['Sporting', 'Sporting Lisbon'],
    rennes: ['Stade Rennais', 'Rennes'],
    brest: ['Stade Brestois', 'Brest'],
    marseille: ['Marseille', 'OM'],
    lyon: ['Olympique Lyonnais', 'Lyon', 'OL'],
    dortmund: ['Borussia Dortmund', 'BVB', 'Dortmund'],
    leverkusen: ['Bayer Leverkusen', 'Leverkusen'],
    cologne: ['FC Köln', 'Koln', 'Cologne'],
    napoli: ['Napoli', 'SSC Napoli'],
    lazio: ['Lazio', 'SS Lazio'],
    roma: ['Roma', 'AS Roma'],
    juventus: ['Juventus', 'Juve'],
    ac_milan: ['AC Milan', 'Milan'],
    south_korea: ['South Korea', 'Korea Republic', 'Republic of Korea'],
    usa: ['USA', 'United States', 'US'],
    england: ['England'],
    netherlands: ['Netherlands', 'Holland'],
    ivory_coast: ["Ivory Coast", "Côte d'Ivoire"],
  };

  for (const v of (knownVariants[id] || [])) variants.add(v);

  return [...variants].filter(Boolean);
}

// ── 主逻辑 ──

const CSV = path.join(__dirname, 'football_teams.csv');
const teamsMap = loadTeams(CSV);
const teams = Object.values(teamsMap);

const issues = [];
const ok = [];

for (const team of teams) {
  const variants = generatePolymarketVariants(team);
  const failing = variants.filter(v => {
    const found = resolve(v, teamsMap);
    return !found || found.id !== team.id;
  });

  if (failing.length > 0) {
    issues.push({ team, failing });
  } else {
    ok.push(team);
  }
}

console.log(`\n✅ 匹配正常（${ok.length} 支）：所有测试变体均可识别\n`);

if (issues.length === 0) {
  console.log('🎉 所有球队均无匹配问题！');
} else {
  console.log(`⚠️  发现 ${issues.length} 支球队存在潜在匹配问题：\n`);
  for (const { team, failing } of issues) {
    console.log(`  ❌ [${team.id}] ${team['zh-CN']} (en: "${team.en}")`);
    for (const f of failing) {
      console.log(`       → Polymarket 可能写法 "${f}" → 无法匹配`);
    }
  }
  console.log('\n建议在 generate.js 的 manualAliases 中补充上述别名。');
}
