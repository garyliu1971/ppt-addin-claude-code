# Claude Code for PowerPoint — POC 实现总结与路线图

## 项目概述

一个 Office JS PowerPoint Add-in，模拟 "Claude Code" 风格，支持自然语言（中/英）通过 AI 操作 PowerPoint 幻灯片、形状、图表、表格、母版、布局和主题。

---

## 实现步骤

### Phase 1 — 基础框架搭建

| # | 步骤 | 文件 |
|---|------|------|
| 1 | 创建 Office Add-in 项目骨架 | `manifest.xml`, `package.json`, `tsconfig.json`, `webpack.config.js` |
| 2 | 任务窗格 HTML/CSS | `src/taskpane.html`, `src/taskpane.css` |
| 3 | Office JS 初始化 + 连接状态 | `src/taskpane.ts` → `init()` |
| 4 | 核心 API 封装层 | `src/services/pptApi.ts` — `runPPT()`, `getSlides()`, `getSelectedSlide()`, `getShapesOnSlide()` |

### Phase 2 — 形状 & 图表 & 表格

| # | 步骤 | 文件 |
|---|------|------|
| 5 | 形状 CRUD（增删改、填充、缩放） | `src/services/shapeService.ts` |
| 6 | 图表 + 表格操作（含降级 fallback） | `src/services/chartTableService.ts` |
| 7 | 原生 API 不可用时自动降级（PPT 2019 兼容） | try/catch → 形状模拟图表/表格 |

### Phase 3 — 母版 & 布局 & 主题 & Slide 操作

| # | 步骤 | 文件 |
|---|------|------|
| 8 | 母版/布局列表、应用布局 | `src/services/masterLayoutThemeService.ts` |
| 9 | Slide 背景、主题读取 | 同上 |
| 10 | Slide CRUD：增删移复制、设标题（含 `slides.add()` 返回 void 的兼容处理） | 同上 |

### Phase 4 — 事件监听

| # | 步骤 | 文件 |
|---|------|------|
| 11 | Shape 选中变化 → 日志输出 | `src/services/eventService.ts` |
| 12 | Slide 切换 → 更新 `currentSlideId` + 日志 | 同上 |
| 13 | 命令执行前二次确认 active slide | `taskpane.ts` → `executeCommand()` 开头 |

### Phase 5 — Regex 命令解析（离线模式）

| # | 步骤 | 文件 |
|---|------|------|
| 14 | 英文正则命令（~15 条） | `taskpane.ts` → `executeCommand()` |
| 15 | 中文正则命令（~10 条）含中文数字（一~十） | 同上 |
| 16 | Quick Action 按钮（9 个） | `taskpane.html` + `taskpane.ts` |

### Phase 6 — AI 集成（DeepSeek API）

| # | 步骤 | 文件 |
|---|------|------|
| 17 | AI Service：18 个 Tool 定义（JSON Schema） | `src/services/aiService.ts` |
| 18 | DeepSeek API 调用（OpenAI 兼容格式） | 同上 → `runAIConversation()` |
| 19 | Function Calling 执行器（switch 路由） | 同上 → `executeToolCall()` |
| 20 | 多轮对话循环（最多 6 轮） | 同上 → for loop + context refresh |
| 21 | API Key 持久化（localStorage）+ UI 状态指示 | 同上 + `taskpane.html/css/ts` |

### Phase 7 — 联网搜索

| # | 步骤 | 文件 |
|---|------|------|
| 22 | CoinGecko API（加密货币，CORS ✅） | `aiService.ts` → `searchWeb()` |
| 23 | Open-Meteo API（天气，CORS ✅） | 同上 |
| 24 | Wikipedia API（百科，CORS ✅） | 同上 |
| 25 | DuckDuckGo + CORS 代理（通用兜底） | 同上 |

---

## 架构总览

```
src/
├── taskpane.html          # 任务窗格 UI
├── taskpane.css           # 样式（紫色渐变主题）
├── taskpane.ts            # 主逻辑：命令分发、AI 集成、UI 事件
├── commands.ts/html       # 功能区命令（stub）
└── services/
    ├── pptApi.ts          # PowerPoint API 封装层
    ├── shapeService.ts    # 形状 CRUD
    ├── chartTableService.ts # 图表 + 表格（含降级）
    ├── masterLayoutThemeService.ts # 母版/布局/主题/Slide CRUD
    ├── eventService.ts    # 事件监听（形状选中、Slide 切换）
    └── aiService.ts       # AI 核心：18 个 Tool、多轮对话、联网搜索
```

**数据流**：
```
用户输入 → taskpane.ts
  ├─ AI 模式: → aiService.runAIConversation()
  │               ├─ 构建上下文 → DeepSeek API
  │               ├─ 解析 tool_calls → executeToolCall()
  │               ├─ 多轮循环（最多 6 轮）
  │               └─ web_search() → CoinGecko/Wikipedia/Open-Meteo
  └─ Regex 模式: → 正则匹配 → 直接调用 Service
                      ↓
              PowerPoint.run() → 文档更新
```

---

## 18 个 AI Tool 矩阵

| 层级 | Tool | 能力 |
|------|------|------|
| **Shape** | `add_shape` | 添加几何形状（15+ 种） |
| | `add_text_box` | 添加文本框 |
| | `set_shape_fill` | 按名改填充色 |
| | `modify_all_shapes` | 批量改样式 |
| | `delete_shape` | 按名删除形状 |
| **Chart/Table** | `add_chart` | 添加柱/饼/折线/环形图 |
| | `add_table` | 添加数据表格 |
| **Slide** | `add_slide` | 新建空白页 |
| | `add_slide_with_title` | 新建带标题页 |
| | `delete_slide_by_index` | 按序号删页 |
| | `set_slide_title` | 设/改标题 |
| | `move_slide` | 移动页位置 |
| | `duplicate_slide` | 复制当前页 |
| **Layout** | `apply_layout` | 应用布局 |
| | `set_slide_background` | 设背景色 |
| **Info** | `list_slides` | 列出所有页 |
| **Search** | `web_search` | 联网搜索（加密币/天气/百科/通用） |
| **Fallback** | `no_op` | 无法执行时说明原因 |

---

## 未来 Plan / 路线图

### 短期（可快速实现）

- [ ] **AI 对话历史** — 保持多轮对话上下文，让 AI 记住之前的操作
- [ ] **文本格式化** — font size/color/bold/italic 在 Tool 中支持
- [ ] **图片插入** — `add_image` Tool（本地上传或 URL）
- [ ] **动画/过渡** — slide transitions, shape animations
- [ ] **导出** — 导出为 PDF/图片
- [ ] **批注/备注** — speaker notes 读写

### 中期（需要一定工作量）

- [ ] **多模型支持** — 切换 OpenAI / Claude / 本地模型
- [ ] **流式响应** — AI 回复逐字显示（SSE streaming）
- [ ] **撤销/重做** — 操作历史栈
- [ ] **模板系统** — 预置 slide 模板（封面、目录、图表页等）
- [ ] **批量生成** — 从 JSON/CSV 数据批量创建 slides
- [ ] **语音输入** — Web Speech API 语音转命令

### 长期 / 产品化

- [ ] **VSCode 集成** — 作为 VSCode 扩展，在编辑器中操作 PPT
- [ ] **协作模式** — 多人同时编辑同一 PPT
- [ ] **Plugin 系统** — 第三方可扩展 Tool
- [ ] **本地模型** — 完全离线运行（Ollama / llama.cpp）
- [ ] **PPT → 网页** — 一键将 PPT 导出为交互式网页
- [ ] **智能排版** — AI 自动优化 slide 布局和美观度

---

## 技术债务 & 已知限制

| 问题 | 说明 | 优先级 |
|------|------|--------|
| CORS 搜索 | DuckDuckGo 被 CORS 拦截，依赖代理/专用 API | 中 |
| `slides.add()` | 返回 void，需 reload 获取新 slide | 已 workaround |
| 图表降级 | 不支持原生 `addChart` 时用矩形模拟 | 已 workaround |
| Office 2019 | 仅支持 PowerPointApi 1.1，多项 API 不可用 | 已降级兼容 |
| API Key 安全 | 存 localStorage，生产环境需后端代理 | 低（POC） |

---

## 本地开发

```bash
npm install
npm run cert          # 安装开发证书
npm run build         # 构建
npm run sideload      # 在 PowerPoint 中启动
# 或手动: npx webpack serve --mode development
```

## Git

- **Repo**: `gary-seismic`
- **Branch**: `personal/gary/AddTaskPane`
- **Commit**: `2db977e` — "feat: full POC"
