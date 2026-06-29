# AI Tool for PowerPoint — POC 构建进度

> Office JS PowerPoint Add-in | TypeScript + Webpack | Task Pane App

## 项目概述

模拟 Claude Code for PowerPoint 概念的 AI 插件，支持自然语言操控 PowerPoint，具备 AI 对话 + 确认即执行工作流。

---

## 技术栈

| 项目 | 值 |
|------|-----|
| 框架 | Office JS PowerPoint API + TypeScript + Webpack 5 |
| AI | DeepSeek Chat API (OpenAI-compatible, Function Calling) |
| 开发服务器 | webpack-dev-server @ `localhost:3000` (HTTPS) |
| Node.js | v24.18.0 |
| Git Branch | `personal/gary/AddTaskPane` |
| 仓库 | `gary-seismic/ppt-addin-claude-code` |

---

## 架构

```
src/
├── taskpane.ts              ← 主 UI 逻辑 (~49KB)
├── taskpane.html            ← 任务窗格 HTML
├── taskpane.css             ← 样式
├── commands.ts              ← Office 命令入口
├── services/
│   ├── pptApi.ts            ← 核心 PowerPoint API 包装 (~3KB)
│   ├── shapeService.ts      ← Shape CRUD + 文字样式 (~9KB)
│   ├── chartTableService.ts ← Chart/Table (~11KB)
│   ├── masterLayoutThemeService.ts ← Slide/Layout/Theme (~12KB)
│   ├── eventService.ts      ← Shape/Slide 事件监听 (~4KB)
│   └── aiService.ts         ← AI 核心 (~32KB)
└── manifest.xml             ← Office Add-in 清单
```

**工作流**: `用户输入 → AI dry-run 规划工具调用 → 展示计划 → 用户确认 → batch 执行 → 自动清理空 slide`

---

## AI 工具列表 (22 个)

### 幻灯片操作
| 工具 | 说明 |
|------|------|
| `add_slide` | 添加空白幻灯片 |
| `add_slide_with_title` | 添加带标题的幻灯片 |
| `delete_slide_by_index` | 删除指定页 (1-based) |
| `set_slide_title` | 设置当前页标题 |
| `move_slide` | 移动幻灯片位置 |
| `duplicate_slide` | 复制当前页 |
| `list_slides` | 列出所有幻灯片 |
| `set_slide_background` | 设置背景色 |

### 形状
| 工具 | 说明 |
|------|------|
| `add_shape` | 添加几何形状 (Rectangle/Oval/Triangle/Star5/...) |
| `add_text_box` | 添加文本框 (支持 left/top/width/height/fontSize) |
| `add_rich_text` | 添加富文本 (每段落独立 fontSize/bold/fontColor) |
| `set_shape_fill` | 按名称填充颜色 |
| `delete_shape` | 按名称删除 |
| `modify_all_shapes` | 批量改样式 |
| `auto_layout` | 自动网格排列 |

### 图表/表格
| 工具 | 说明 |
|------|------|
| `add_table` | 插入/更新表格 (upsert) |
| `add_chart` | 插入图表 (Column/Bar/Pie/Line/Area) |

### 主题/设计
| 工具 | 说明 |
|------|------|
| `apply_layout` | 应用布局 |
| `apply_theme` | 应用主题 (22 个内置) |
| `apply_design_scheme` | 应用配色方案 (7 个) |
| `list_themes` | 列出主题和方案 |
| `web_search` | 网络搜索 (CoinGecko/Wikipedia/DuckDuckGo) |

---

## 已完成的 Regex 命令 (AI-off 模式回落)

| 类别 | 命令 | 中文 |
|------|------|------|
| 形状 | `add rectangle` / `add oval` / ... | `添加矩形` / `添加圆形` |
| | `add text box "xxx"` | `添加文本框 "xxx"` |
| | `make all shapes blue` | `所有形状设为蓝色` |
| | `fill shape "name" with red` | `填充形状 "xxx" 红色` |
| | `resize shape 200x150` | — |
| | `delete shape "name"` | — |
| 文字 | `font size 10` | `字体大小 10` |
| | `bold text` | `粗体` |
| | `set text color red` | `文字颜色 红` |
| 图表 | `add bar chart` / `add pie chart` | `添加柱状图` |
| | `add table` | `添加表格` |
| 幻灯片 | `add slide titled "xxx"` | `添加幻灯片 标题为 "xxx"` |
| | `delete slide 3` | `删掉第3页` |
| | `set title to "xxx"` | `设置标题为 "xxx"` |
| | `move slide 1 to 3` | `移动第1页到第3页` |
| | `duplicate slide` | `复制当前页` |
| | `set background blue` | `设置背景 蓝` |
| | `apply layout "xxx"` | — |
| | `apply theme xxx` | `应用主题 xxx` |
| 查询 | `list slides` / `list shapes` / `list layouts` | `列出幻灯片` / `列出形状` |
| | `page setup` / `slide size` | `页面设置` / `幻灯片尺寸` |
| | `list themes` | `列出主题` |
| | `apply design xxx` | `应用设计 xxx` |

---

## 关键设计决策

1. **DryRun → Confirm → Execute**: AI 先规划工具调用，用户预览后一键 Apply
2. **自动清理空 slide**: 执行后删除只含 0-1 个 shape 的**新建** slide（不动已有 slide）
3. **对话记忆**: localStorage 保存最近 16 条消息，跨命令保持上下文
4. **深度搜索**: CoinGecko (币价) → Wikipedia → DuckDuckGo (最多 2 次)
5. **自动换 target slide**: `add_slide` 后自动将后续工具指向新 slide

---

## 已知限制

| 限制 | 说明 |
|------|------|
| Office 2019 | `addChart`/`addTable`/`slideWidth` 等 API 不可用 |
| 段落 API | `paragraphs` 在 TypeScript 类型中未声明，需 `as any` |
| 多 slide 操作 | AI 需手动指定 `target_slide`，否则默认当前页 |
| 粗体/字号混排 | `add_rich_text` 可做到，`add_text_box` 全 text 同一样式 |

---

## 未来计划 (PLAN.md)

- **Phase 8**: 更多工具 (图片、视频、动画)
- **Phase 9**: 上下文记忆增强
- **Phase 10**: 专业样式模板库

---

## 开发命令

```bash
# 启动开发服务器
npm start

# 构建
npm run build

# Git (当前分支)
git branch: personal/gary/AddTaskPane
```

## 最近提交

1. `15d7dd7` — add_rich_text: 段落级格式化
2. `2a1af62` — 只删新建空 slide，保护已有 slide
3. `4af838b` — 自动删除空 slide
4. `172e62a` — AI 必须生成内容，禁止空壳 slide
5. `3620f3a` — 精确定位 + 字体控制
6. `70dcc8c` — 修复 Send 按钮 + add 形状不需要 "shape" 关键字
