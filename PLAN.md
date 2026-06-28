# AI Tool for PowerPoint — 实现总结 & 改进路线图

## 目标

模拟 Claude Code for PowerPoint — Chat 方式编辑 PPT，支持中/英自然语言。

---

## 当前能力（22 个 Tool，POC v1）

| 层级 | Tools | 状态 |
|------|-------|:--:|
| Shape | add_shape, add_text_box, set_shape_fill, modify_all_shapes, delete_shape | ✅ |
| Chart/Table | add_chart, add_table (upsert) | ✅ |
| Slide CRUD | add_slide, delete_slide_by_index, set_slide_title, move_slide, duplicate_slide | ✅ |
| Design | apply_layout, set_slide_background, apply_theme, apply_design_scheme, auto_layout | ✅ |
| Info | list_slides, list_themes, web_search | ✅ |

**架构**：DryRun → Preview → Apply | target_slide | Conversation History | Slide Index Context

---

## 三大改进方向

### 1. 更多 Tool Executor

| 优先级 | Tool | 能力 |
|:--:|------|------|
| 🔴 | `add_image` | 插入图片 |
| 🔴 | `align_shapes` / `distribute_shapes` | 对齐/分布 |
| 🔴 | `add_speaker_notes` | 演讲备注 |
| 🟡 | `group_shapes` / `ungroup` | 组合 |
| 🟡 | `set_animation` / `set_transition` | 动画/切换 |
| 🟡 | `add_smartart` | SmartArt |
| 🟢 | `import_data` (Excel/CSV→slides) | 数据导入 |
| 🟢 | `export_slide` (PNG/PDF) | 导出 |

### 2. 增强 Context & Memory

| 改进 | 说明 | 优先级 |
|------|------|:--:|
| **User Profile** | 记住用户语言、主题偏好、常用颜色 | 🔴 |
| **Slide Structure** | 注入占位符类型、布局信息 | 🔴 |
| **Summary Compression** | 长对话自动摘要，保留关键决策 | 🟡 |
| **Design Rules** | System prompt 注入排版原则（6×6、对比、留白） | 🔴 |
| **Tool Result Memory** | 重要操作（slide ID, table shape）短期记忆 | 🟡 |

### 3. Professional PPTX 样式

| 能力 | 说明 |
|------|------|
| **Design Rules in Prompt** | 6×6 规则、字号阶梯、配色限制、留白要求 |
| **Smart Templates** | title_slide / content_bullet / two_column / chart_page / closing_slide |
| **Color Advisor** | 按行业自动推荐配色（科技蓝、金融绿...） |
| **Auto Layout** | 根据内容类型选择最佳模板 |
| **Font Pairing** | 标题+正文字体组合 |

---

## 路线图

### Phase 8（当前）— 交互体验
- [ ] 10+ 新 Tool（图片、对齐、SmartArt、动画、备注）
- [ ] User Profile Memory
- [ ] Design Rules System Prompt
- [ ] 5 个预置 slide 模板

### Phase 9 — 深度整合
- [ ] Summary Compression
- [ ] Streaming 响应
- [ ] Undo/Redo
- [ ] 批量数据导入

### Phase 10 — 产品化
- [ ] 多模型（OpenAI/Claude/Ollama）
- [ ] 对话导出
- [ ] 离线模式

---

## Git

| Commit | Message |
|--------|---------|
| `8ceacf3` | confirm-before-execute, target_slide, auto-layout, upsert table, slide index context, rebrand |
| `2db977e` | AI multi-turn, Chinese commands, web search, slide ops |
| `5827a25` | Initial POC |

**Branch**: `personal/gary/AddTaskPane` | **Repo**: `gary-seismic`
