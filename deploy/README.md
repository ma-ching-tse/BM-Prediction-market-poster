# 部署配置存档

线上服务器（阿里云轻量香港，47.79.20.248）的关键运维文件，方便重建/迁移时复用。

## 文件清单

| 文件 | 服务器路径 | 作用 |
|---|---|---|
| `poster.service` | `/etc/systemd/system/poster.service` | systemd 守护进程定义，让 `proxy.js` 开机自启、崩溃自动重启 |
| `poster-cleanup.cron` | `/etc/cron.d/poster-cleanup` | 每天 03:00 清理 `output/` 下超过 7 天的产物 |

## 服务器初始部署速记

```bash
# 1. 装基础依赖
apt-get update && apt-get install -y zip unzip fonts-noto-cjk
curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
apt-get install -y nodejs

# 2. 拉代码
mkdir -p /opt/poster && cd /opt/poster
# rsync 项目过来（含 lark.config.json，注意 .gitignore 不会带过去，要单独同步）
npm install   # 会自动下载 Puppeteer 自带 Chromium

# 3. 装 systemd 和 cron
cp deploy/poster.service /etc/systemd/system/
cp deploy/poster-cleanup.cron /etc/cron.d/poster-cleanup
chmod 644 /etc/cron.d/poster-cleanup
systemctl daemon-reload
systemctl enable --now poster.service
```

## 注意事项

- `lark.config.json` 是 `.gitignore` 的，重建服务器时要**单独**从开发机同步过去
- 服务器目前用 `root` 直接跑（轻量服务器、内部使用，简化处理）；密码请定期更换
- 公网端口 3002 在阿里云控制台防火墙放行
