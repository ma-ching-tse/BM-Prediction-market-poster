# Classic 模板更新手册（NBA 赛事）

## 1. 模板定位

- 模板名：`classic`
- 模板文件：`poster.html`
- 输出目录：`output/NBA/<日期>/`

## 2. 数据来源与字段

Classic 现在从 Lark 读取“主文案 + 3 场比赛链接”的模板数据（配置见 `lark.config.json` 的 `spreadsheetToken`、`sheetId`）。

表头至少要有：

1. `lang`
2. `title`
3. `subtitle`
4. `footer`
5. `match1_link`
6. `match2_link`
7. `match3_link`

可选字段：

1. `match1_home` / `match1_away` / `match1_date`
2. `match1_home_win` / `match1_away_win`
3. `match2_home` / `match2_away` / `match2_date`
4. `match2_home_win` / `match2_away_win`
5. `match3_home` / `match3_away` / `match3_date`
6. `match3_home_win` / `match3_away_win`

说明：

1. 源语言行通常是 `zh-CN`，翻译结果会从第 3 行开始回填。
2. 现在按 `match1_link ~ match3_link` 直接贴 3 场 Polymarket 对应场次链接。
3. 脚本会自动解析并回填主队、客队、日期和赔率。

## 3. 更新方式 A：按 Polymarket 链接更新（最常用）

操作：在 Lark 的源语言行填写：

1. `title`
2. `subtitle`
3. `footer`
4. `match1_link`
5. `match2_link`
6. `match3_link`

脚本行为：

1. 按 `matchN_link` 读取 Polymarket 比赛信息
2. 自动解析并回填 `matchN_home / away / date / home_win / away_win`
3. 自动翻译 `title / subtitle / footer`
4. 自动把翻译结果写回第 3 行开始
5. 自动生成多语言海报

执行命令：

```bash
npm run generate
```

当前出图规格：

1. 输出格式：PNG
2. 输出尺寸：`900x900`
3. 压缩方式：`ffmpeg` 调色板压缩
4. 目标体积：单张约 `300KB` 内

## 4. 更新方式 B：改多语言文案/字号

编辑文件：`poster.copy.json` 下的 `classic` 节点。

这里主要保留字号类参数，不再作为 NBA 主文案来源。

你可以调整：

1. `titleFontSize`
2. `titleLineHeight`
3. `titleMaxWidth`
4. `subtitleFontSize`
5. `subtitleLineHeight`

建议：先改单语种，执行一次看效果，再批量改其他语种。

## 5. 更新方式 C：改球队名或 Logo

1. `teams.csv`：维护球队 ID 与多语言名、logo 名称
2. `NBA_icon/`：放置 logo 文件，文件名需与 `teams.csv` 的 `logo` 列一致

例如 `logo=洛杉矶湖人`，则必须有 `NBA_icon/洛杉矶湖人.png`。

## 6. 更新方式 D：改背景或版式

1. 多语言背景图：`backgrounds/<lang>.png`
2. 基础素材：`assets/vs_new.png`、`assets/track_new.png`
3. 结构样式：`poster.html`

注意：脚本按 `backgrounds/` 下的文件名识别语言并决定输出语种。

## 7. 运行结果检查

重点检查：

1. 是否生成 `output/NBA/<YYYYMMDD>/NBA_<YYYYMMDD>.zip`
2. 每种语言是否都有 PNG
3. 表格里的 `matchN_home / away / date / home_win / away_win` 是否已自动回填
4. 标题是否溢出、球队 logo 是否缺失
5. 终端是否提示 `已回填 Lark 表格`，且没有出现 Polymarket 解析失败报错
6. 单张图片尺寸是否为 `900x900`，体积是否大致在 `300KB` 内
