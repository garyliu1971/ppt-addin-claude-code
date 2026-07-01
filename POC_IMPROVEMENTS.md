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
