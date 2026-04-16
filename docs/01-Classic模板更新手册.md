# Classic 模板更新手册（NBA 赛事）

## 1. 模板定位

- 模板名：`classic`
- 模板文件：`poster.html`
- 输出目录：`output/NBA/<日期>/`

## 2. 数据来源与字段

Classic 从 Lark 普通电子表格读取比赛数据（配置见 `lark.config.json` 的 `spreadsheetToken`、`sheetId`、`range`）。

表头至少要有：

1. `date`
2. `home_team`
3. `away_team`

可选字段：

1. `home_win`
2. `away_win`
3. `polymarket_slug`
4. `home_outcome`
5. `away_outcome`

## 3. 更新方式 A：只改比赛数据（最常用）

操作：在 Lark 改比赛日期和主客队 ID。

脚本行为：

1. 自动拉取 Polymarket 概率
2. 自动回填 `home_win` / `away_win` 到 Lark
3. 自动生成多语言海报

执行命令：

```bash
npm run generate
```

## 4. 更新方式 B：人工锁定赔率

如果你希望使用人工赔率，不走自动抓取：

1. 在 Lark 中手动填 `home_win` 和 `away_win`
2. 确保两者加总等于 `100`

脚本会做校验，若不等于 100 会失败。

## 5. 更新方式 C：修正 Polymarket 对应关系

若自动匹配不到正确市场：

1. 在 Lark 填 `polymarket_slug`（手动指定市场）
2. 必要时填 `home_outcome` / `away_outcome`（手动指定 outcome）

这样可以避免映射错误。

## 6. 更新方式 D：改多语言文案/字号

编辑文件：`poster.copy.json` 下的 `classic` 节点。

你可以调整：

1. `title`
2. `subtitle`
3. `titleFontSize`
4. `titleLineHeight`
5. `titleMaxWidth`
6. `subtitleFontSize`
7. `subtitleLineHeight`

建议：先改单语种，执行一次看效果，再批量改其他语种。

## 7. 更新方式 E：改球队名或 Logo

1. `teams.csv`：维护球队 ID 与多语言名、logo 名称
2. `NBA_icon/`：放置 logo 文件，文件名需与 `teams.csv` 的 `logo` 列一致

例如 `logo=洛杉矶湖人`，则必须有 `NBA_icon/洛杉矶湖人.png`。

## 8. 更新方式 F：改背景或版式

1. 多语言背景图：`backgrounds/<lang>.png`
2. 基础素材：`assets/vs_new.png`、`assets/track_new.png`
3. 结构样式：`poster.html`

注意：脚本按 `backgrounds/` 下的文件名识别语言并决定输出语种。

## 9. 运行结果检查

重点检查：

1. 是否生成 `output/NBA/<YYYYMMDD>/NBA_<YYYYMMDD>.zip`
2. 每种语言是否都有 JPG
3. 标题是否溢出、球队 logo 是否缺失
4. 终端是否提示 `已回填 Lark 表格`，且没有出现胜率校验失败报错
