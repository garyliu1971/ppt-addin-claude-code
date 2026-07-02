/**
 * aiService.ts — AI-powered command processing
 * Multi-turn conversation with DeepSeek + web search via DuckDuckGo.
 */

import {
  getSlides, getSelectedSlide, getShapesOnSlide, getPresentationInfo,
} from "./pptApi";
import {
  addShape, addTextBox, addImage, setShapeFill, deleteShape, applyStyleToAllShapes,
  detectOverlaps, autoLayoutShapes, addStructuredTextBox,
  addImageFromBase64, setShapeFormat, addCard,
} from "./shapeService";
import { addTable, addChart, upsertTable, ChartType, ChartData } from "./chartTableService";
import {
  applyLayoutToSlide, findLayoutByName, setSlideBackground,
  addSlide, deleteSlide, deleteSlideByIndex, setSlideTitle,
  moveSlide, duplicateSlide, addSlideWithTitle, getSlidesWithIndex,
  applyTheme, listAvailableThemes, applyDesignScheme, listDesignSchemes,
} from "./masterLayoutThemeService";
import { buildProfessionalSlide, getNBADemoData } from "./slideBuilderService";
import { validateToolCall, validateBatch, readBackQA } from "./validateService";

// ── Config ────────────────────────────────────────────────────────

const STORAGE_KEY = "ppt_ai_api_key";
let apiKey = "";
try { const s = localStorage.getItem(STORAGE_KEY); if (s) apiKey = s; } catch { /* */ }

const API_BASE = "https://api.deepseek.com/v1";
const MODEL = "deepseek-chat";

// ── Conversation History (persists across commands) ───────────────

let conversationHistory: any[] = [];

export function clearConversationHistory(): void {
  conversationHistory = [];
}

export function getHistoryLength(): number {
  return conversationHistory.length;
}

export function setApiKey(k: string): void { apiKey = k; try { localStorage.setItem(STORAGE_KEY, k); } catch { /* */ } }
export function clearApiKey(): void { apiKey = ""; try { localStorage.removeItem(STORAGE_KEY); } catch { /* */ } }
export function getApiKey(): string { return apiKey; }
export function hasApiKey(): boolean { return apiKey.length > 0; }

// ── Web Search (multi-source: CoinGecko, Wikipedia, DuckDuckGo) ─────

async function searchWeb(query: string): Promise<string> {
  const q = query.toLowerCase();

  // ── Crypto prices: CoinGecko (free, CORS-enabled) ──────────────
  if (/\b(bitcoin|btc|crypto|ethereum|eth|doge|dogecoin|solana|sol|ripple|xrp|cardano|ada)\b/i.test(q)) {
    const coinMap: Record<string, string> = {
      bitcoin: "bitcoin", btc: "bitcoin", ethereum: "ethereum", eth: "ethereum",
      doge: "dogecoin", dogecoin: "dogecoin", solana: "solana", sol: "solana",
      xrp: "ripple", ripple: "ripple", cardano: "cardano", ada: "cardano",
    };
    let coinId = "bitcoin";
    for (const [kw, id] of Object.entries(coinMap)) {
      if (q.includes(kw)) { coinId = id; break; }
    }

    try {
      const url = `https://api.coingecko.com/api/v3/simple/price?ids=${coinId}&vs_currencies=usd&include_24hr_change=true`;
      const r = await fetch(url, { signal: AbortSignal.timeout(5000) });
      if (r.ok) {
        const d = await r.json();
        if (d[coinId]) {
          const price = d[coinId].usd;
          const change = d[coinId].usd_24h_change;
          const changeStr = change != null ? ` (24h: ${change > 0 ? "+" : ""}${change.toFixed(2)}%)` : "";
          return `💰 ${coinId.toUpperCase()} price: $${price?.toLocaleString() || price} USD${changeStr} [CoinGecko]`;
        }
      }
    } catch { /* try next source */ }
  }

  // ── Weather: Open-Meteo (free, CORS-enabled) ────────────────────
  if (/weather|天气|temperature|forecast|climate/i.test(q)) {
    try {
      // Default to a location — user can refine in prompt
      const url = `https://api.open-meteo.com/v1/forecast?latitude=39.9&longitude=116.4&current_weather=true`;
      const r = await fetch(url, { signal: AbortSignal.timeout(5000) });
      if (r.ok) {
        const d = await r.json();
        if (d.current_weather) {
          const w = d.current_weather;
          return `🌤 Weather: ${w.temperature}°C, wind ${w.windspeed} km/h, code ${w.weathercode} [Open-Meteo]`;
        }
      }
    } catch { /* try next source */ }
  }

  // ── Wikipedia (search API → find page → get summary) ──────────
  try {
    // Step 1: Search Wikipedia for the query
    const searchUrl = `https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch=${encodeURIComponent(query)}&format=json&origin=*&srlimit=3`;
    const searchResp = await fetch(searchUrl, { signal: AbortSignal.timeout(5000) });
    if (searchResp.ok) {
      const searchData = await searchResp.json();
      const results = searchData?.query?.search;
      if (results?.length) {
        // Step 2: Get summary for the top result
        const pageTitle = encodeURIComponent(results[0].title);
        const summaryUrl = `https://en.wikipedia.org/api/rest_v1/page/summary/${pageTitle}`;
        const summaryResp = await fetch(summaryUrl, { signal: AbortSignal.timeout(5000) });
        if (summaryResp.ok) {
          const d = await summaryResp.json();
          if (d.extract) return `📚 ${d.title}: ${d.extract.slice(0, 800)}... [Wikipedia]`;
        }
        // Fallback: use search snippet
        return `📚 Wikipedia: ${results.slice(0, 3).map((r: any) => `• ${r.title}: ${r.snippet.replace(/<[^>]+>/g, "")}`).join("\n")}`;
      }
    }
  } catch { /* try next source */ }

  // ── DuckDuckGo instant answer (still works for simple queries) ──
  try {
    const ddgUrl = `https://api.duckduckgo.com/?q=${encodeURIComponent(query)}&format=json&no_html=1&skip_disambig=1`;
    const r = await fetch(ddgUrl, { signal: AbortSignal.timeout(5000) });
    if (r.ok) {
      const d = await r.json();
      if (d.AbstractText) return `📚 ${d.AbstractText.slice(0, 500)} [DuckDuckGo]`;
      if (d.Answer) return `📚 ${d.Answer} [DuckDuckGo]`;
      if (d.RelatedTopics?.length) {
        return d.RelatedTopics.slice(0, 5)
          .filter((t: any) => t.Text)
          .map((t: any) => `• ${t.Text}`)
          .join("\n");
      }
    }
  } catch { /* try next source */ }

  return "⚠️ Web search returned no results. Use your own knowledge about this topic to create the slide content. Do NOT search again — proceed with what you know.";
}

// ── Tool Definitions ──────────────────────────────────────────────

interface ToolDef {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, any> };
}

const TOOLS: ToolDef[] = [
  { type: "function", function: { name: "add_shape", description: "Add a geometric shape. Types: Rectangle, RoundRectangle, Oval/Ellipse, Triangle, Diamond, Arrow, Star5, Heart, Cloud, Sun, Moon, SmileyFace. For numbered badges or labels INSIDE shapes, use text+textColor+fontSize (text auto-centers).", parameters: { type: "object", properties: { geometry: { type: "string" }, fillColor: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, transparency: { type: "number" }, lineVisible: { type: "boolean" }, lineColor: { type: "string" }, lineWeight: { type: "number" }, rotation: { type: "number" }, text: { type: "string", description: "Text to place inside the shape (auto-centered, e.g. badge number)" }, textColor: { type: "string" }, fontSize: { type: "number" } }, required: ["geometry"] } } },
  { type: "function", function: { name: "add_image", description: "Add an image from a URL to the current slide. Use full image URLs (jpg, png, svg, gif). For base64 data URIs, use add_image_base64.", parameters: { type: "object", properties: { url: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" } }, required: ["url"] } } },
  { type: "function", function: { name: "add_image_base64", description: "Add an image from a base64 data URI directly. Copilot: slide.shapes.addImage(base64, {left, top, width, height}).", parameters: { type: "object", properties: { base64: { type: "string", description: "Full base64 data URI (data:image/png;base64,...)" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, transparency: { type: "number" } }, required: ["base64"] } } },
  { type: "function", function: { name: "add_card", description: "Create a content card WITH text in one call. Automatically creates RoundRectangle background + heading text box + optional subtitle. Use for match brackets, info cards, list items. heading=team name/score, subtitle=date/venue.", parameters: { type: "object", properties: { left: { type: "number" }, top: { type: "number" }, width: { type: "number", description: "default 430" }, height: { type: "number", description: "default 52" }, fillColor: { type: "string" }, lineColor: { type: "string" }, heading: { type: "string", description: "Main text (team name, score, title)" }, headingSize: { type: "number" }, headingColor: { type: "string" }, subtitle: { type: "string", description: "Secondary text (date, venue, details)" }, subtitleSize: { type: "number" }, subtitleColor: { type: "string" } }, required: ["left", "top", "heading"] } } },
  { type: "function", function: { name: "add_text_box", description: "Add a text box with full positioning. Supports alignment (Left/Center/Right/Justify), verticalAlignment (Top/Middle/Bottom), margins (left/right/top/bottom in pt). Slide is typically 960x540pt (widescreen).", parameters: { type: "object", properties: { text: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, fontSize: { type: "number" }, horizontalAlignment: { type: "string" }, verticalAlignment: { type: "string" }, leftMargin: { type: "number" }, rightMargin: { type: "number" }, topMargin: { type: "number" }, bottomMargin: { type: "number" } }, required: ["text"] } } },
  { type: "function", function: { name: "add_rich_text", description: "Add a text box with PER-PARAGRAPH formatting. Each paragraph can have fontSize, bold, fontColor, italic, alignment. Also supports box-level verticalAlignment and margins. Use for structured content.", parameters: { type: "object", properties: { left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, verticalAlignment: { type: "string" }, leftMargin: { type: "number" }, rightMargin: { type: "number" }, topMargin: { type: "number" }, bottomMargin: { type: "number" }, paragraphs: { type: "array", items: { type: "object", properties: { text: { type: "string" }, fontSize: { type: "number" }, bold: { type: "boolean" }, italic: { type: "boolean" }, fontColor: { type: "string" }, alignment: { type: "string" } }, required: ["text"] } } }, required: ["paragraphs"] } } },
  { type: "function", function: { name: "modify_all_shapes", description: "Apply a style to all shapes on current slide. Supports fillColor, fontSize, transparency, lineVisible.", parameters: { type: "object", properties: { fillColor: { type: "string" }, fontSize: { type: "number" }, transparency: { type: "number" }, lineVisible: { type: "boolean" } } } } },
  { type: "function", function: { name: "set_shape_fill", description: "Set fill color of a specific shape by name.", parameters: { type: "object", properties: { shapeName: { type: "string" }, color: { type: "string" } }, required: ["shapeName", "color"] } } },
  { type: "function", function: { name: "delete_shape", description: "Delete a shape by name.", parameters: { type: "object", properties: { shapeName: { type: "string" } }, required: ["shapeName"] } } },
  { type: "function", function: { name: "add_table", description: "Insert a data table.", parameters: { type: "object", properties: { headers: { type: "array", items: { type: "string" } }, rows: { type: "array", items: { type: "array", items: { type: "string" } } } }, required: ["headers", "rows"] } } },
  { type: "function", function: { name: "add_chart", description: "Insert a chart. chartType: ColumnClustered|BarClustered|Pie|Line|Doughnut|Area", parameters: { type: "object", properties: { chartType: { type: "string" }, title: { type: "string" }, categories: { type: "array", items: { type: "string" } }, series: { type: "array", items: { type: "object", properties: { name: { type: "string" }, values: { type: "array", items: { type: "number" } } } } } }, required: ["chartType", "categories", "series"] } } },
  { type: "function", function: { name: "set_slide_background", description: "Set slide background color.", parameters: { type: "object", properties: { color: { type: "string" } }, required: ["color"] } } },
  { type: "function", function: { name: "apply_layout", description: "Apply a slide layout by name.", parameters: { type: "object", properties: { layoutName: { type: "string" } }, required: ["layoutName"] } } },
  { type: "function", function: { name: "add_slide", description: "Add a new blank slide.", parameters: { type: "object", properties: {} } } },
  { type: "function", function: { name: "add_slide_with_title", description: "Add a new slide with a title.", parameters: { type: "object", properties: { title: { type: "string" } }, required: ["title"] } } },
  { type: "function", function: { name: "delete_slide_by_index", description: "Delete a slide by number (1-based).", parameters: { type: "object", properties: { index: { type: "number" } }, required: ["index"] } } },
  { type: "function", function: { name: "set_slide_title", description: "Set the title of the current slide.", parameters: { type: "object", properties: { title: { type: "string" } }, required: ["title"] } } },
  { type: "function", function: { name: "move_slide", description: "Move a slide to a new position.", parameters: { type: "object", properties: { fromIndex: { type: "number" }, toIndex: { type: "number" } }, required: ["fromIndex", "toIndex"] } } },
  { type: "function", function: { name: "duplicate_slide", description: "Duplicate the current slide.", parameters: { type: "object", properties: {} } } },
  { type: "function", function: { name: "list_slides", description: "List all slides.", parameters: { type: "object", properties: {} } } },
  { type: "function", function: { name: "apply_theme", description: "Apply a PowerPoint theme by name. Available themes: Office, Facet, Ion, Organic, Slice, Wisp, Integral, Retrospect, Parallax, Celestial, Gallery, Mesh, Savon, Berlin, Crop, Circuit, Depth, Droplet, Headlines, Metropolitan, View, Wood Type.", parameters: { type: "object", properties: { themeName: { type: "string" } }, required: ["themeName"] } } },
  { type: "function", function: { name: "apply_design_scheme", description: "Apply a designer color scheme to the current slide. Available: modern dark, ocean blue, forest green, sunset orange, rose gold, clean white, slate gray.", parameters: { type: "object", properties: { schemeName: { type: "string" } }, required: ["schemeName"] } } },
  { type: "function", function: { name: "list_themes", description: "List all available PowerPoint themes and design schemes.", parameters: { type: "object", properties: {} } } },
  { type: "function", function: { name: "web_search", description: "Search the web for real-time data (sports scores, weather, news, stock prices). Use when you need current information.", parameters: { type: "object", properties: { query: { type: "string" } }, required: ["query"] } } },
  { type: "function", function: { name: "auto_layout", description: "Auto-arrange ALL shapes on the current slide into a neat grid. Detects overlaps and repositions. Use when shapes overlap or user asks to rearrange/format/align shapes.", parameters: { type: "object", properties: { columns: { type: "number", description: "Number of columns (default 3)" } } } } },
  { type: "function", function: { name: "set_shape_format", description: "Apply Copilot formatting to a specific shape: fill, line, font, alignment, margins, transparency. Copilot: shape.fill.setSolidColor, shape.fill.transparency, shape.lineFormat.visible/color/weight, textFrame.verticalAlignment, textFrame.leftMargin/rightMargin, paragraphFormat.horizontalAlignment.", parameters: { type: "object", properties: { shapeName: { type: "string" }, fillColor: { type: "string" }, transparency: { type: "number" }, lineColor: { type: "string" }, lineWeight: { type: "number" }, lineVisible: { type: "boolean" }, rotation: { type: "number" }, fontSize: { type: "number" }, bold: { type: "boolean" }, italic: { type: "boolean" }, fontName: { type: "string" }, fontColor: { type: "string" }, alignment: { type: "string" }, verticalAlignment: { type: "string" }, leftMargin: { type: "number" }, rightMargin: { type: "number" }, topMargin: { type: "number" }, bottomMargin: { type: "number" } }, required: ["shapeName"] } } },
  { type: "function", function: { name: "no_op", description: "Use when request cannot be fulfilled.", parameters: { type: "object", properties: { message: { type: "string" } }, required: ["message"] } } },
  { type: "function", function: { name: "build_professional_slide", description: "Build a complete professional slide in ONE call with background, title, description, multi-column cards, insight bar, and footer. Use when user wants a 'professional', 'magazine-style', 'data-driven', or 'NBA/FIFA-style' slide. Accepts a full JSON schema. For NBA demo, pass nba_demo=true.", parameters: { type: "object", properties: { nba_demo: { type: "boolean", description: "Set to true to build the built-in NBA demo slide" }, data: { type: "object", description: "Full ProfessionalSlideSchema JSON. See slideBuilderService.ts for the schema shape." } } } } },
];

// ── Context Builder ───────────────────────────────────────────────

async function buildSlideContext(): Promise<string> {
  try {
    const info = await getPresentationInfo();
    const slide = await getSelectedSlide();
    let ctx = `Presentation: "${info.title}" | ${info.slideCount} slide(s)\n`;
    if (slide) {
      // Find the 1-based index of the current slide
      const slides = await getSlides();
      const idx = slides.findIndex(s => s.id === slide.id) + 1;
      ctx += `Current slide: Slide ${idx || "?"} (id: ${slide.id})\n`;
      const shapes = await getShapesOnSlide(slide.id);
      if (shapes.length > 0) {
        ctx += `Shapes (${shapes.length}) — existing positions:\n`;
        for (const s of shapes) {
          const r = Math.round(s.left + s.width);
          const b = Math.round(s.top + s.height);
          ctx += `  - "${s.name || "?"}" type=${s.type} pos=(${Math.round(s.left)},${Math.round(s.top)})→(${r},${b}) ${Math.round(s.width)}x${Math.round(s.height)}\n`;
        }
        const maxBottom = Math.max(...shapes.map(s => Math.round(s.top + s.height)));
        ctx += `  Lowest shape bottom edge: y=${maxBottom}\n`;
      }
    }
    return ctx;
  } catch { return "Context unavailable."; }
}

// ── Multi-turn Conversation ───────────────────────────────────────

export interface ExecResult { success: boolean; message: string; }
export interface PendingCall { name: string; args: Record<string, any>; description: string; }

export interface AIResult {
  messages: string[];
  toolResults: ExecResult[];
  pendingCalls: PendingCall[];
}

export async function runAIConversation(
  userCommand: string,
  currentSlideId: string | null,
  dryRun: boolean = false
): Promise<AIResult> {
  if (!apiKey) throw new Error("API key not set.");

  const sysPrompt = `You are a PowerPoint AI assistant. Your job: create professional slides with REAL content.

Today: ${new Date().toISOString().split("T")[0]}. Slide context: ${await buildSlideContext()}

CRITICAL — follow these rules:
1. ALWAYS add content to slides. When creating a slide, include add_text_box / add_card / add_rich_text with REAL substantive text in the SAME turn. A slide with only a title is a FAILURE.
1b. Unless user explicitly asks to create a NEW slide, modify the CURRENT slide only. Do NOT call add_slide unless asked.
2. When user asks for introductions, details, bios, history, or facts about ANY topic — ALWAYS call web_search first. Use the real facts returned. web_search max 3 times.
3. Plan ALL elements for one slide in ONE turn. Don't split slide creation across multiple turns.
4. Multi-slide: every write call must include "target_slide" (1-based index).
5. Slide dimensions: 960×540pt. Use precise left/top/width/height values.
6. Colors: blue=#4A90D9, red=#E74C3C, green=#2ECC71, yellow=#F1C40F, orange=#E67E22, purple=#9B59B6, black=#333333, white=#FFFFFF, dark navy=#1a1a2e.
7. For cards/lists: use add_card (creates roundrect + text in one call). For text: add_text_box or add_rich_text. For shapes: add_shape.
8. Layout patterns:
   - Simple slide: add_slide → set_slide_background → add_text_box (title, top=20) → add_text_box (body, top=80)
   - Cards grid: add_card with col=i%2, row=floor(i/2), w=430, h=52, pitch=58
   - Bracket: add_card for every match (heading=teams, subtitle=date/venue), 2-column grid
9. If the request is vague, create ONE well-designed slide with substantive content. Better one good slide than multiple empty ones.
10. When done, confirm what was created.
11. **Professional slides**: when the user wants a polished, data-driven slide with cards/columns/insight bars — use build_professional_slide with the full JSON schema. This creates background, eyebrow, title, description, column cards, insight footer all at once. For NBA/football team slides, pass nba_demo=true.

─── FEW-SHOT EXAMPLES ───
Example 1: User says "add a blue rectangle"
→ add_shape(geometry="Rectangle", fillColor="#4A90D9", left=150, top=150, width=200, height=120)

Example 2: User says "add a title and body text on a dark slide"
→ set_slide_background(color="#1a1a2e")
→ add_text_box(text="Quarterly Report", left=60, top=50, width=840, height=60, fontSize=32)
→ add_text_box(text="Revenue grew 15% YoY driven by...", left=60, top=130, width=840, height=300, fontSize=16)

Example 3: User says "create 3 cards in a row"
→ add_card(left=40, top=100, width=280, height=80, heading="Card 1", subtitle="Details here")
→ add_card(left=340, top=100, width=280, height=80, heading="Card 2", subtitle="Details here")
→ add_card(left=640, top=100, width=280, height=80, heading="Card 3", subtitle="Details here")
‼️ CRITICAL: When user asks for N cards, you MUST call add_card N times — one call per card. Do NOT put all content in a single card.

─── CARD LAYOUT RULES ───
• 2 cards: cols at left=40 and left=490, width=430
• 3 cards: cols at left=40, left=330, left=620, width=280
• 4 cards: cols at left=20, left=255, left=490, left=725, width=215
• When adding cards AFTER existing shapes: start at top = max(existing shapes bottom edge) + 20
• NEVER place cards at top < 140 when a title exists — cards will overlap title text
• Vertical spacing: card rows at pitch = height + 10

─── VALIDATION RULES (enforced) ───
• All shapes within slide (960x540). left+width ≤ 960, top+height ≤ 540.
• Min shape size: 2pt. Colors: hex #RRGGBB or named (blue, red, etc).
• Overlap >30% between planned shapes is BLOCKED — use correct positions.`;

  const textMessages: string[] = [];
  const toolResults: ExecResult[] = [];
  const pendingCalls: PendingCall[] = [];
  let searchCount = 0;
  let movedSlides = new Set<string>();
  const addedSlideTitles = new Set<string>(); // prevent duplicate slide creation
  const MAX = 8;

  // Build messages: always include current context
  const messages: any[] = [];
  if (conversationHistory.length === 0) {
    messages.push({ role: "system", content: sysPrompt });
  } else {
    // Inject current context as a system note so AI doesn't re-ask what's on the slide
    messages.push({ role: "system", content: `Current context: ${await buildSlideContext()}Stay on current slide unless asked to add a new one.` });
  }
  messages.push(...conversationHistory);
  messages.push({ role: "user", content: userCommand });

  // Trim history to keep context manageable (keep last 12 messages = ~6 turns)
  if (conversationHistory.length > 12) {
    conversationHistory = conversationHistory.slice(-8);
  }

  for (let turn = 0; turn < MAX; turn++) {
    const resp = await fetch(`${API_BASE}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
      body: JSON.stringify({ model: MODEL, messages, tools: TOOLS, tool_choice: "auto", temperature: 0.5, max_tokens: 8192 }),
    });

    if (!resp.ok) throw new Error(`API error ${resp.status}: ${await resp.text()}`);
    const data = await resp.json();
    const msg = data.choices?.[0]?.message;
    if (!msg) break;

    const assistantMsg: any = { role: "assistant", content: msg.content || "" };
    if (msg.tool_calls?.length) assistantMsg.tool_calls = msg.tool_calls;
    messages.push(assistantMsg);

    if (msg.content) textMessages.push(msg.content);
    if (!msg.tool_calls?.length) break; // AI is done

    for (const tc of msg.tool_calls) {
      let args: any = {};
      try { args = JSON.parse(tc.function.arguments); } catch { /* */ }

      let result: ExecResult;
      const isReadOnly = tc.function.name === "list_slides" || tc.function.name === "list_themes" ||
                         tc.function.name === "web_search";

      // ── Step 5: Validate & Repair (pre-execution) ──
      if (!isReadOnly) {
        const validation = validateToolCall(tc.function.name, args);
        if (!validation.valid) {
          const errMsg = validation.errors.map(e => `❌ ${e.field || "?"}: ${e.message}`).join("; ");
          result = { success: false, message: `Validation failed: ${errMsg}` };
        } else if (validation.warnings.length > 0) {
          console.warn(`[validate] ${tc.function.name} warnings:`, validation.warnings.map(w => w.message));
        }
      }

      // ── Dedup: block duplicate add_slide AFTER validation ──
      if (!result) {
        const isSlideAdd = tc.function.name === "add_slide" || tc.function.name === "add_slide_with_title";
        if (isSlideAdd && addedSlideTitles.has(args.title || "")) {
          result = { success: false, message: `⚠️ Duplicate slide "${args.title || "untitled"}" skipped.` };
        } else {
          // Track as used (even in dryRun, to prevent duplicates in planned actions)
          if (isSlideAdd) addedSlideTitles.add(args.title || "");

          // In dryRun mode, still execute read-only tools so AI gets real data
          if (dryRun && !isReadOnly) {
            result = { success: true, message: `[Preview] Would call ${tc.function.name}` };
          } else if (tc.function.name === "web_search" && searchCount >= 3) {
            result = { success: false, message: "⚠️ Maximum 3 web searches reached. Create slides NOW with what you know." };
          } else if (tc.function.name === "move_slide" && movedSlides.has(String(args.fromIndex))) {
            result = { success: false, message: `⚠️ Slide ${args.fromIndex} was already moved — cannot move again.` };
          } else {
            if (tc.function.name === "web_search") searchCount++;
            if (tc.function.name === "move_slide") movedSlides.add(String(args.fromIndex));
            result = await executeToolCall(tc.function.name, args, currentSlideId);
          }
        }
      }

      // Only push to pendingCalls if not a duplicate
      if (result.success || !result.message.includes("skipped")) {
        toolResults.push(result);
        pendingCalls.push({ name: tc.function.name, args, description: `${tc.function.name}(${JSON.stringify(args).slice(0, 200)})` });
      } else {
        toolResults.push(result);
      }

      messages.push({
        role: "tool", tool_call_id: tc.id,
        content: (dryRun && !isReadOnly)
          ? JSON.stringify({ success: true, message: `[Preview] Would execute ${tc.function.name}` })
          : JSON.stringify({ success: result.success, message: result.message, context: dryRun ? "" : await buildSlideContext() }),
      });
    }

    // ── Batch overlap check after all tool calls in this turn ──
    if (dryRun && pendingCalls.length > 1) {
      const batchResults = validateBatch(
        pendingCalls.map(c => ({ name: c.name, args: c.args }))
      );
      for (const br of batchResults) {
        if (br.errors.length > 0) {
          toolResults.push({
            success: false,
            message: `⚠️ Overlap: ${br.errors.map(e => e.message).join("; ")}`,
          });
        }
      }
    }
  }

  // Save to persistent history — skip in dryRun mode
  if (!dryRun) {
    conversationHistory = messages
      .filter((m: any) => m.role === "user" || m.role === "assistant")
      .map((m: any) => {
        if (m.role === "assistant" && m.tool_calls) {
          const { tool_calls, ...rest } = m;
          return rest;
        }
        return m;
      })
      .slice(-16);
  }

  return { messages: textMessages, toolResults, pendingCalls };
}

// ── Tool Executor ─────────────────────────────────────────────────

/** Execute a batch of pending calls (called after user confirms) */
export async function executePendingCalls(
  calls: PendingCall[],
  sid: string | null
): Promise<ExecResult[]> {
  const results: ExecResult[] = [];
  let defaultSid = sid;
  let slideCount = 0;

  // Get initial slide count for tracking new slides
  try { slideCount = (await getSlides()).length; } catch { /* */ }

  // Remember pre-existing slide IDs so we only clean up NEW empty slides
  const preExistingIds = new Set<string>();
  try {
    const slides = await getSlides();
    for (const s of slides) preExistingIds.add(s.id);
  } catch { /* */ }

  for (const call of calls) {
    if (call.name === "list_slides" || call.name === "list_themes" || call.name === "no_op" || call.name === "web_search") {
      results.push({ success: true, message: `[Skipped] ${call.description}` });
      continue;
    }

    // Resolve target slide: explicit target_slide arg > defaultSid
    const targetSid = call.args.target_slide
      ? await resolveSlideId(call.args.target_slide, defaultSid)
      : defaultSid;

    console.log(`[execute] ${call.name} → slide=${targetSid || "(none)"}`, call.args);
    const r = await executeToolCall(call.name, call.args, targetSid);
    console.log(`[execute] ${call.name} → ${r.success ? "OK" : "FAIL"}: ${r.message}`);
    results.push(r);

    // After adding a slide, auto-track it as default for subsequent tools
    if (r.success && (call.name === "add_slide" || call.name === "add_slide_with_title")) {
      try {
        // Extract slide ID from the success message
        const idMatch = r.message.match(/id:\s*([^\s)]+)/);
        if (idMatch) {
          defaultSid = idMatch[1];
          slideCount++;
        } else {
          // Fallback: reload slides
          const slides = await getSlides();
          if (slides.length > slideCount) {
            defaultSid = slides[slides.length - 1].id;
            slideCount = slides.length;
          }
        }
      } catch { /* */ }
    }
  }

  // ── Post-processing: delete NEW empty slides (don't touch pre-existing) ──
  try {
    const allSlides = await getSlides();
    for (const slide of allSlides) {
      // Only check slides created during this batch
      if (preExistingIds.has(slide.id)) continue;
      try {
        const shapes = await getShapesOnSlide(slide.id);
        // Slide is empty if it has ≤1 shapes (only the default title placeholder)
        if (shapes.length <= 1) {
          await deleteSlide(slide.id);
          results.push({ success: true, message: `🗑 Removed empty new slide` });
        }
      } catch { /* skip if can't check */ }
    }
  } catch { /* skip cleanup */ }

  // ── Step 6: Read-back QA (post-execution verification) ──
  if (defaultSid) {
    try {
      const writeCalls = calls.filter(c =>
        !["list_slides", "list_themes", "no_op", "web_search", "apply_theme"].includes(c.name)
      );
      if (writeCalls.length > 0) {
        const qa = await readBackQA(defaultSid, 1, writeCalls.map(c => c.name));
        results.push({ success: qa.ok, message: qa.message });
      }
    } catch { /* QA is best-effort */ }
  }

  return results;
}

const CM: Record<string, string> = { blue: "#4A90D9", red: "#E74C3C", green: "#2ECC71", yellow: "#F1C40F", orange: "#E67E22", purple: "#9B59B6", pink: "#E91E63", black: "#333333", white: "#FFFFFF", gray: "#95A5A6" };
const toHex = (c: string) => CM[c.toLowerCase()] || c;

// ── Slide Resolver ────────────────────────────────────────────────

/** Resolve a 1-based slide index to a slide ID. Returns null for "current". */
async function resolveSlideId(index: number | undefined, defaultSid: string | null): Promise<string | null> {
  if (index === undefined || index === null) return defaultSid;
  try {
    const slides = await getSlides();
    const target = slides[index - 1];
    if (target) return target.id;
  } catch { /* fall through */ }
  return defaultSid; // fallback to current if resolution fails
}

/** Extract target_slide from args and resolve to slide ID */
async function getTargetSid(args: Record<string, any>, defaultSid: string | null): Promise<string | null> {
  return resolveSlideId(args.target_slide as number | undefined, defaultSid);
}

export async function executeToolCall(
  name: string, args: Record<string, any>, sid: string | null
): Promise<ExecResult> {
  try {
    switch (name) {
      case "add_shape":
        await addShape(args.geometry, { fillColor: args.fillColor ? toHex(args.fillColor) : "#4A90D9", left: args.left, top: args.top, width: args.width, height: args.height, transparency: args.transparency, lineVisible: args.lineVisible, lineColor: args.lineColor ? toHex(args.lineColor) : undefined, lineWeight: args.lineWeight, rotation: args.rotation, text: args.text, textColor: args.textColor ? toHex(args.textColor) : undefined, fontSize: args.fontSize }, sid ?? undefined);
        return { success: true, message: `Added ${args.geometry}` };

      case "add_image":
        await addImage(args.url, { left: args.left, top: args.top, width: args.width, height: args.height }, sid ?? undefined);
        return { success: true, message: `Added image from ${args.url}` };

      case "add_image_base64":
        await addImageFromBase64(args.base64, { left: args.left, top: args.top, width: args.width, height: args.height, transparency: args.transparency }, sid ?? undefined);
        return { success: true, message: `Added image from base64` };

      case "add_card":
        await addCard({
          left: args.left ?? 40, top: args.top ?? 80,
          width: args.width ?? 430, height: args.height ?? 52,
          fillColor: args.fillColor ? toHex(args.fillColor) : undefined,
          lineColor: args.lineColor ? toHex(args.lineColor) : undefined,
          heading: args.heading || "Untitled",
          headingSize: args.headingSize,
          headingColor: args.headingColor ? toHex(args.headingColor) : undefined,
          subtitle: args.subtitle,
          subtitleSize: args.subtitleSize,
          subtitleColor: args.subtitleColor ? toHex(args.subtitleColor) : undefined,
        }, sid ?? undefined);
        return { success: true, message: `Added card: "${args.heading || "Untitled"}"` };

      case "add_text_box":
        await addTextBox(args.text, { left: args.left, top: args.top, width: args.width, height: args.height, fontSize: args.fontSize, horizontalAlignment: args.horizontalAlignment, verticalAlignment: args.verticalAlignment, leftMargin: args.leftMargin, rightMargin: args.rightMargin, topMargin: args.topMargin, bottomMargin: args.bottomMargin }, sid ?? undefined);
        return { success: true, message: `Added text: "${args.text.slice(0, 60)}${args.text.length > 60 ? "..." : ""}"` };

      case "add_rich_text": {
        const paras = args.paragraphs;
        if (!paras || !Array.isArray(paras) || paras.length === 0) {
          return { success: false, message: "add_rich_text requires 'paragraphs' array" };
        }
        await addStructuredTextBox(paras, { left: args.left, top: args.top, width: args.width, height: args.height, verticalAlignment: args.verticalAlignment, leftMargin: args.leftMargin, rightMargin: args.rightMargin, topMargin: args.topMargin, bottomMargin: args.bottomMargin }, sid ?? undefined);
        return { success: true, message: `Added rich text: ${paras.length} paragraph(s)` };
      }

      case "modify_all_shapes": {
        if (!sid) return { success: false, message: "No slide selected" };
        const n = await applyStyleToAllShapes(sid, { fillColor: args.fillColor ? toHex(args.fillColor) : undefined, fontSize: args.fontSize, transparency: args.transparency, lineVisible: args.lineVisible });
        return { success: true, message: `Styled ${n} shape(s)` };
      }

      case "set_shape_fill": {
        if (!sid) return { success: false, message: "No slide selected" };
        const shapes = await getShapesOnSlide(sid);
        const t = shapes.find(s => (s.name || "").toLowerCase().includes(args.shapeName.toLowerCase()));
        if (!t) return { success: false, message: `Shape "${args.shapeName}" not found` };
        await setShapeFill(t.id, sid, toHex(args.color));
        return { success: true, message: `Filled "${t.name}"` };
      }

      case "delete_shape": {
        if (!sid) return { success: false, message: "No slide selected" };
        const shapes = await getShapesOnSlide(sid);
        const t = shapes.find(s => (s.name || "").toLowerCase().includes(args.shapeName.toLowerCase()));
        if (!t) return { success: false, message: `Shape "${args.shapeName}" not found` };
        await deleteShape(t.id, sid);
        return { success: true, message: `Deleted "${t.name}"` };
      }

      case "add_table":
        var tr = await upsertTable({ headers: args.headers, rows: args.rows }, sid ?? undefined);
        return { success: true, message: tr };

      case "add_chart":
        await addChart(args.chartType as ChartType, { categories: args.categories, series: args.series }, { title: args.title }, sid ?? undefined);
        return { success: true, message: `Added ${args.chartType} chart` };

      case "set_slide_background": {
        if (!sid) return { success: false, message: "No slide selected" };
        await setSlideBackground(sid, toHex(args.color));
        return { success: true, message: `Background set to ${args.color}` };
      }

      case "apply_layout": {
        if (!sid) return { success: false, message: "No slide selected" };
        const f = await findLayoutByName(args.layoutName);
        if (!f) return { success: false, message: `Layout "${args.layoutName}" not found` };
        await applyLayoutToSlide(sid, f.layoutId);
        return { success: true, message: `Applied layout "${args.layoutName}"` };
      }

      case "add_slide": { const s = await addSlide(); return { success: true, message: `Added slide (id: ${s?.id || "?"})` }; }
      case "add_slide_with_title": { await addSlideWithTitle(args.title); return { success: true, message: `Added slide: "${args.title}"` }; }
      case "delete_slide_by_index": { const m = await deleteSlideByIndex(args.index); return { success: true, message: m }; }
      case "set_slide_title": { if (!sid) return { success: false, message: "No slide selected" }; await setSlideTitle(sid, args.title); return { success: true, message: `Title: "${args.title}"` }; }

      case "move_slide": {
        const slides = await getSlidesWithIndex();
        const from = slides.find(s => s.index === args.fromIndex);
        if (!from) return { success: false, message: `Slide ${args.fromIndex} not found` };
        await moveSlide(from.id, args.toIndex);
        return { success: true, message: `Moved ${args.fromIndex}→${args.toIndex}` };
      }

      case "duplicate_slide": {
        if (!sid) return { success: false, message: "No slide selected" };
        await duplicateSlide(sid);
        return { success: true, message: "Duplicated slide" };
      }

      case "list_slides": {
        const slides = await getSlides();
        return { success: true, message: slides.map((s, i) => `Slide ${i + 1}: ${s.id}`).join("\n") };
      }

      case "web_search": {
        const r = await searchWeb(args.query);
        return { success: true, message: `🔍 "${args.query}":\n${r}` };
      }

      case "auto_layout": {
        if (!sid) return { success: false, message: "No slide selected" };
        const overlaps = await detectOverlaps(sid);
        const n = await autoLayoutShapes(sid, args.columns || 3);
        const msg = overlaps.length > 0
          ? `Rearranged ${n} shapes (fixed ${overlaps.length} overlap${overlaps.length > 1 ? "s" : ""})`
          : `Arranged ${n} shapes into grid`;
        return { success: true, message: msg };
      }

      case "apply_theme": {
        const msg = await applyTheme(args.themeName);
        return { success: true, message: msg };
      }

      case "apply_design_scheme": {
        if (!sid) return { success: false, message: "No slide selected" };
        const msg = await applyDesignScheme(sid, args.schemeName);
        return { success: true, message: msg };
      }

      case "list_themes": {
        const themes = listAvailableThemes();
        const schemes = listDesignSchemes();
        return { success: true, message: `Themes (${themes.length}): ${themes.join(", ")}\nDesign Schemes (${schemes.length}): ${schemes.join(", ")}` };
      }

      case "set_shape_format": {
        if (!sid) return { success: false, message: "No slide selected" };
        const shapes = await getShapesOnSlide(sid);
        const t = shapes.find(s => (s.name || "").toLowerCase().includes(args.shapeName.toLowerCase()));
        if (!t) return { success: false, message: `Shape "${args.shapeName}" not found` };
        await setShapeFormat(t.id, sid, {
          fillColor: args.fillColor ? toHex(args.fillColor) : undefined,
          transparency: args.transparency,
          lineColor: args.lineColor ? toHex(args.lineColor) : undefined,
          lineWeight: args.lineWeight,
          lineVisible: args.lineVisible,
          rotation: args.rotation,
          fontSize: args.fontSize,
          bold: args.bold,
          italic: args.italic,
          fontName: args.fontName,
          fontColor: args.fontColor ? toHex(args.fontColor) : undefined,
          alignment: args.alignment,
          verticalAlignment: args.verticalAlignment,
          leftMargin: args.leftMargin,
          rightMargin: args.rightMargin,
          topMargin: args.topMargin,
          bottomMargin: args.bottomMargin,
        });
        return { success: true, message: `Formatted "${t.name}"` };
      }

      case "build_professional_slide": {
        if (args.nba_demo) {
          const r = await buildProfessionalSlide(getNBADemoData(), sid ?? undefined);
          return { success: true, message: `Built NBA demo slide: ${r.shapeCount} shapes on slide ${r.slideId}` };
        }
        if (args.data) {
          const r = await buildProfessionalSlide(args.data, sid ?? undefined);
          return { success: true, message: `Built professional slide: ${r.shapeCount} shapes on slide ${r.slideId}` };
        }
        return { success: false, message: "build_professional_slide requires 'data' schema or 'nba_demo=true'" };
      }

      case "no_op":
        return { success: false, message: args.message || "Cannot fulfill" };

      default:
        return { success: false, message: `Unknown tool: ${name}` };
    }
  } catch (err: any) {
    return { success: false, message: `${name}: ${err.message || err}` };
  }
}
