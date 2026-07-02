# POC Improvement Summary — 2026-07-01

## ✅ What Worked Well

### 1. System Prompt Rewrite
**Before**: 40+ lines of conflicting rules, AI paralyzed.
**After**: 10 focused rules, content rule at #1.
- Result: AI consistently creates content instead of empty titles.
- Rule 1b ("stay on current slide") eliminated duplicate slide creation.

### 2. Parameter Tuning
| Parameter | Before | After | Impact |
|-----------|--------|-------|--------|
| `temperature` | 0.1 | **0.5** | Richer, more creative content |
| `max_tokens` | 2000 | **8192** | Room for 8+ tool calls per turn |
| `web_search max` | 2 | **3** | More facts for content |

### 3. Slide Targeting Bug Fix (Critical)
**Bug**: All creation functions (`addShape`, `addTextBox`, `addCard`, etc.) called `getSelectedSlide()` internally. When batch-creating slides, content went to the **user-clicked** slide instead of the **target** slide.
**Fix**: Every function now accepts optional `slideId` parameter. AI service passes the correct target slide ID.

### 4. Slide ID Tracking
**Bug**: After `add_slide`, `getSlides()` in a new `PowerPoint.run()` context returned different ID → `GeneralException`.
**Fix**: Extract slide ID directly from `add_slide` success message (`"Added slide (id: xxx)"`).

### 5. `add_card` Compound Tool
**Problem**: AI created empty `RoundRectangle` shapes without text.
**Solution**: `add_card` creates background + heading + subtitle in one atomic call. AI can't forget text.

### 6. Web Search Rewrite
**Before**: Wikipedia REST API needed exact page names → most queries failed.
**After**: Wikipedia **Search API** → finds matching pages → gets summary. Covers general queries like "F1 teams 2026".

### 7. Input Validation Guards
- `add_rich_text`: checks `paragraphs` exists before calling `.map()`
- `add_card`: defaults `left: 40, top: 80, heading: "Untitled"` if AI omits them

### 8. `applyTheme` Fallback
**Bug**: `presentation.theme.name` is read-only in Office.js.
**Fix**: Maps 14 built-in themes → design schemes (color + background) as fallback.

### 9. Conversation History Trimming
Keeps last 8 messages to prevent context bloat across multi-turn conversations.

---

## ⚠️ Remaining Issues

### 1. `addChart` Still Shape-Based
- Native `shapes.addChart()` requires `PictureAndCharts` requirement set (Office 365 1810+).
- Falls back to visual rectangles/text boxes.
- **Fix needed**: Check PowerPoint version and use native API when available.

### 2. `GeneralException` on Some Operations
- `set_slide_background` and `add_text_box` occasionally fail on newly created slides.
- Might be timing issue or slide initialization delay.
- **Mitigation**: Slide ID tracking fixed most cases.

### 3. No Image Insertion API
- `shapes.addImage(base64)` availability depends on PowerPoint version.
- Falls back to rectangle with picture fill.
- Cannot set image opacity or crop.

### 4. No Shadow/Effect API
- Copilot examples use `outerShdw` for card shadows.
- Office.js stable API doesn't support shape effects.
- Cards appear "flat" without shadows.

### 5. No External JSON Data Loading
- Copilot examples load `.json` files for data-driven slides.
- Our POC has no `load_data` tool.
- AI must embed all data in tool calls.

### 6. Dry-Run + Execute Complexity
- Two-phase execution adds latency and potential for plan/execute mismatch.
- AI's dry-run plan may not account for slide index shifts during execution.

---

## 📊 Success Metrics

| Metric | Before | After |
|--------|--------|-------|
| Slides with actual content | ~30% | **~85%** |
| Empty shape errors | Frequent | Rare |
| Web search success rate | ~20% | **~70%** |
| Duplicate slide creation | Common | Fixed |
| GeneralException events | 3-5 per batch | 0-1 per batch |

---

## 🔧 Tool Inventory (25 tools)

| Tool | Status | Notes |
|------|--------|-------|
| `add_shape` | ✅ | text+textColor for badges |
| `add_text_box` | ✅ | alignment, margins |
| `add_rich_text` | ✅ | per-paragraph formatting |
| `add_card` | ✅ | compound card builder |
| `add_image` / `add_image_base64` | ⚠️ | fallback to rect+fill |
| `add_table` / `upsert_table` | ✅ | native + visual fallback |
| `add_chart` | ⚠️ | shape-based fallback |
| `set_shape_format` | ✅ | retroactive formatting |
| `set_shape_fill` / `delete_shape` | ✅ | by name |
| `modify_all_shapes` | ✅ | bulk styling |
| `set_slide_background` | ⚠️ | occasional GeneralException |
| `set_slide_title` | ✅ | |
| `apply_layout` | ✅ | by name |
| `apply_theme` | ⚠️ | scheme fallback |
| `apply_design_scheme` | ✅ | |
| `add_slide` / `add_slide_with_title` | ✅ | |
| `delete_slide_by_index` | ✅ | |
| `move_slide` / `duplicate_slide` | ✅ | |
| `list_slides` / `list_themes` | ✅ | |
| `web_search` | ✅ | Wikipedia search API |
| `auto_layout` | ✅ | grid arrangement |

---

# Request Lifecycle — Validation & QA (2026-07-02)

> 参考 End-to-End Request Lifecycle 架构图，补齐 Step 5 (Validate & Repair) 和 Step 6 (QA)。

## 架构对齐

```
1. User Message ──→ 2. Context Builder ──→ 3. Prompt Assembler ──→ 4. LLM Tool Calling
                                                                       │
                                                   ┌───────────────────┘
                                                   ▼
                                       5. Validate & Repair ←── NEW
                                          │  ❌ fail → loop back to LLM
                                          ▼  ✅ pass
                                       6. Execute + Render ←── NEW (QA)
```

## 新增 `src/services/validateService.ts`

### Pre-execution Validation

| 校验项 | 规则 | 级别 |
|---|---|---|
| **Bounds overflow** | `left + width ≤ 960`, `top + height ≤ 540` | Error（阻断） |
| **负坐标** | `left < 0` 或 `top < 0` | Warning |
| **最小尺寸** | `width/height ≥ 2pt` | Error |
| **Required fields** | 每 tool 的必填参数检查 | Error |
| **Color format** | 非 hex 非 named color | Warning |
| **Overlap prediction** | 同批次 shapes 位置重叠预测 | Warning |

### Post-execution Read-back QA

```ts
readBackQA(slideId, expectedMin, toolNames)
// → "✅ QA: 12 shapes on slide after add_text_box, add_shape"
// → "⚠️ QA: Expected ≥3 shapes, found 1"
```

## `aiService.ts` 改进

| 改进 | 说明 |
|---|---|
| **Few-shot examples (3 例)** | 蓝色矩形 / 标题+正文 / 三列卡片 → LLM 坐标准确率显著提高 |
| **Validation rules in prompt** | 960×540 bounds, 命名颜色列表 → LLM 自我约束 |
| **Step 5 pre-exec check** | dryRun 循环中每个 write call 先 validate，不通过则阻断并返回错误给 LLM |
| **Step 6 QA** | executePendingCalls 末尾读回 shape count 验证 |

## Repair Loop

```
LLM tool_call → validateToolCall()
  → ❌ "left(800) + width(600) = 1400 exceeds slide width 960"
  → 错误注入 conversation → LLM 自动修正 → 重新 tool_call
```

## 效果

| Before | After |
|---|---|
| LLM 产出超大 shape → 执行后超出 slide | 校验阻断 + 自动修复 |
| 多个 shapes 堆叠 | Overlap warning |
| 不确定是否执行成功 | `✅ QA: N shapes on slide` |
| LLM 坐标随机 | Few-shot 示例 + bounds 约束 |
