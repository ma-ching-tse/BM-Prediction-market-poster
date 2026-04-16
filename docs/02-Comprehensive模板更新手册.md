# Comprehensive 模板更新手册（综合事件）

## 1. 模板定位

- 模板名：`comprehensive`
- 模板文件：`poster.comprehensive-event.html`
- 输出目录：`output/综合事件/<日期>/`

## 2. 数据来源（Lark）

读取 `lark.config.json` 中 `comprehensiveSheetId` 指向的工作表。

当前脚本固定读取 `A1:J2`：

1. 第 1 行：表头
2. 第 2 行：源语言内容（默认 `sourceLang=zh-CN`）

建议表头使用：

1. `main title`
2. `sub title`
3. `footer`
4. `1`
5. `percent_1`
6. `2`
7. `percent_2`
8. `3`
9. `percent_3`

## 3. 更新方式 A：改活动文案和比例（最常用）

1. 在 Lark 第 2 行修改标题、副标题、底部文案、3 条问题文本和百分比
2. 运行生成命令

```bash
npm run generate:comprehensive
```

脚本会：

1. 读取第 2 行源语言内容
2. 按背景语种自动翻译其他语言（MyMemory）
3. 将翻译结果回填到第 3 行开始（A 列是语言代码）
4. 生成海报与 ZIP

## 4. 更新方式 B：改多语言 UI 文案与字号策略

编辑 `poster.copy.json` 下的 `comprehensive` 节点。

常改字段：

1. `title` / `subtitle` / `footer`
2. `outcomeLabel`
3. `titleFontSize` / `titleLineHeight` / `titleMaxWidth`
4. `subtitleFontSize` / `subtitleLineHeight`
5. `cardTextFontSize` / `cardTextLineHeight`

说明：综合事件里，卡片问题文本优先使用翻译结果；`poster.copy.json` 主要用于样式和兜底文案。

## 5. 更新方式 C：改卡片配图

编辑 `poster.copy.json` 的 `comprehensive.default.cards[i].image`，例如：

- `assets/card-icons/btc.png`
- `assets/card-icons/soccer.png`
- `assets/card-icons/trump.png`

如果不填图，会走兜底逻辑（可能显示为空或使用其他兜底内容）。

## 6. 更新方式 D：控制输出语种

语种由 `backgrounds/` 文件名决定。

例如存在：

- `backgrounds/zh-CN.png`
- `backgrounds/en.png`
- `backgrounds/ja.png`

则会输出这些语种；删除某语种背景即可停产该语种。

## 7. 更新方式 E：改版式与视觉

1. 改结构样式：`poster.comprehensive-event.html`
2. 改背景风格：`backgrounds/<lang>.png`
3. 改品牌 Logo：`poster.copy.json` 的 `brandLogo`

## 8. 运行结果检查

1. 是否生成 `output/综合事件/<YYYYMMDD>/综合事件_<YYYYMMDD>.zip`
2. 第 3 行开始是否成功回填翻译内容
3. 百分比显示是否正确
4. 长标题是否在 2 行内可读
