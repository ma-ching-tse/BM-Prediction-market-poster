#!/bin/bash
# ─────────────────────────────────────────────
#  BitMart 海报生成器 — 首次安装脚本（Mac）
#  运行方式：在终端里执行  bash setup.sh
# ─────────────────────────────────────────────

set -e

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

ok()   { echo -e "${GREEN}✅ $1${NC}"; }
warn() { echo -e "${YELLOW}⚠️  $1${NC}"; }
err()  { echo -e "${RED}❌ $1${NC}"; exit 1; }
step() { echo -e "\n${YELLOW}▶ $1${NC}"; }

echo ""
echo "🎨  BitMart 海报生成器 — 首次安装"
echo "────────────────────────────────"

# ── 1. Homebrew ──────────────────────────────
step "检查 Homebrew..."
if command -v brew &>/dev/null; then
  ok "Homebrew 已安装"
else
  warn "未找到 Homebrew，开始安装（需要输入开机密码）..."
  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" || err "Homebrew 安装失败，请截图联系技术同学"

  # Apple Silicon 需要追加 PATH
  if [[ -f /opt/homebrew/bin/brew ]]; then
    eval "$(/opt/homebrew/bin/brew shellenv)"
    echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> "$HOME/.zprofile"
  fi
  ok "Homebrew 安装完成"
fi

# ── 2. Node.js ───────────────────────────────
step "检查 Node.js..."
if command -v node &>/dev/null; then
  NODE_VER=$(node -v)
  ok "Node.js 已安装 ($NODE_VER)"
else
  warn "未找到 Node.js，开始安装..."
  brew install node || err "Node.js 安装失败，请截图联系技术同学"
  ok "Node.js 安装完成 ($(node -v))"
fi

# ── 3. ffmpeg ────────────────────────────────
step "检查 ffmpeg（图片压缩工具）..."
if command -v ffmpeg &>/dev/null; then
  ok "ffmpeg 已安装"
else
  warn "未找到 ffmpeg，开始安装（约 300MB，需要等几分钟）..."
  brew install ffmpeg || err "ffmpeg 安装失败，请截图联系技术同学"
  ok "ffmpeg 安装完成"
fi

# ── 4. npm install ───────────────────────────
step "安装项目依赖（npm install）..."
npm install || err "依赖安装失败，请截图联系技术同学"
ok "项目依赖安装完成"

# ── 5. 提示配置 lark.config.json ─────────────
echo ""
echo "────────────────────────────────"
echo -e "${GREEN}🎉  安装完成！${NC}"
echo ""
echo "下一步："
echo "  1. 用文本编辑器打开项目文件夹里的 lark.config.json"
echo "  2. 填入 appId 和 appSecret（找技术同学获取）"
echo "  3. 按运营手册第一步启动服务：npm start"
echo ""
