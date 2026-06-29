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
} from "./shapeService";
import { addTable, addChart, upsertTable, ChartType, ChartData } from "./chartTableService";
import {
  applyLayoutToSlide, findLayoutByName, setSlideBackground,
  addSlide, deleteSlide, deleteSlideByIndex, setSlideTitle,
  moveSlide, duplicateSlide, addSlideWithTitle, getSlidesWithIndex,
  applyTheme, listAvailableThemes, applyDesignScheme, listDesignSchemes,
} from "./masterLayoutThemeService";

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

  // ── Wikipedia (CORS-enabled) ───────────────────────────────────
  try {
    const searchTerm = encodeURIComponent(query.replace(/price|price of|what is|how much|today|live|now/gi, "").trim().slice(0, 50));
    const wikiUrl = `https://en.wikipedia.org/api/rest_v1/page/summary/${searchTerm}`;
    const r = await fetch(wikiUrl, { signal: AbortSignal.timeout(5000) });
    if (r.ok) {
      const d = await r.json();
      if (d.extract) return `📚 ${d.title}: ${d.extract.slice(0, 500)}... [Wikipedia]`;
    }
  } catch { /* try next source */ }

  // ── DuckDuckGo via proxy (last resort) ─────────────────────────
  const ddgUrl = `https://api.duckduckgo.com/?q=${encodeURIComponent(query)}&format=json&no_html=1&skip_disambig=1`;
  const proxies = [
    `https://corsproxy.io/?url=${encodeURIComponent(ddgUrl)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(ddgUrl)}`,
  ];

  for (const proxyUrl of proxies) {
    try {
      const r = await fetch(proxyUrl, { signal: AbortSignal.timeout(8000) });
      if (!r.ok) continue;
      const raw = await r.text();
      let data: any;
      try { data = JSON.parse(raw); } catch { continue; }
      // allorigins wraps in contents
      if (data.contents) data = JSON.parse(data.contents);

      if (data.AbstractText) return data.AbstractText;
      if (data.Answer) return data.Answer;
      if (data.RelatedTopics) {
        return data.RelatedTopics.slice(0, 5)
          .filter((t: any) => t.Text)
          .map((t: any) => `• ${t.Text}`)
          .join("\n");
      }
    } catch { /* try next */ }
  }

  return "⚠️ Web search returned no results. Use your own knowledge about this topic to create the slide content. Do NOT search again — proceed with what you know.";
}

// ── Tool Definitions ──────────────────────────────────────────────

interface ToolDef {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, any> };
}

const TOOLS: ToolDef[] = [
  { type: "function", function: { name: "add_shape", description: "Add a geometric shape. Types: Rectangle, Oval, Triangle, Diamond, Arrow, Star5, Heart, Cloud, Sun, Moon, SmileyFace.", parameters: { type: "object", properties: { geometry: { type: "string" }, fillColor: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" } }, required: ["geometry"] } } },
  { type: "function", function: { name: "add_image", description: "Add an image from a URL to the current slide. Use full image URLs (jpg, png, svg, gif).", parameters: { type: "object", properties: { url: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" } }, required: ["url"] } } },
  { type: "function", function: { name: "add_text_box", description: "Add a text box with full positioning. Use for headers, body text, footers. Slide is typically 960x540pt (widescreen) or 720x540pt (4:3).", parameters: { type: "object", properties: { text: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, fontSize: { type: "number" } }, required: ["text"] } } },
  { type: "function", function: { name: "add_rich_text", description: "Add a text box with PER-PARAGRAPH formatting. Each paragraph can have its own fontSize, bold, fontColor. Use for structured content where headings (bold 16pt) and body (normal 10pt) must coexist in one box. Paragraphs array: [{text, fontSize?, bold?, fontColor?}].", parameters: { type: "object", properties: { left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, paragraphs: { type: "array", items: { type: "object", properties: { text: { type: "string" }, fontSize: { type: "number" }, bold: { type: "boolean" }, fontColor: { type: "string" } }, required: ["text"] } } }, required: ["paragraphs"] } } },
  { type: "function", function: { name: "modify_all_shapes", description: "Apply a style to all shapes on current slide.", parameters: { type: "object", properties: { fillColor: { type: "string" }, fontSize: { type: "number" } } } } },
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
  { type: "function", function: { name: "no_op", description: "Use when request cannot be fulfilled.", parameters: { type: "object", properties: { message: { type: "string" } }, required: ["message"] } } },
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
        ctx += `Shapes (${shapes.length}):\n`;
        for (const s of shapes) {
          ctx += `  - "${s.name || "?"}" type=${s.type} pos=(${Math.round(s.left)},${Math.round(s.top)}) ${Math.round(s.width)}x${Math.round(s.height)}\n`;
        }
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

  const sysPrompt = `You are a PowerPoint AI assistant. Use tools to fulfill requests.

Today's date: ${new Date().toISOString().split("T")[0]} (${new Date().toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" })}). Always use this as the current date.

Current context: ${await buildSlideContext()}

Rules:
- CRITICAL: When operating on MULTIPLE slides in one command, you MUST include "target_slide" (1-based) on EVERY write tool call. Example: set_slide_title({"title":"A","target_slide":1}), set_slide_background({"color":"blue","target_slide":2}).
- After add_slide, the NEW slide becomes the default. To target the OLD slide, explicitly set target_slide.
- If unsure which slide is which, call list_slides first.
- Use your own knowledge. Only web_search for explicit "latest/live/current/today/real-time" requests.
- add_table updates existing tables automatically. Do NOT delete tables.
- web_search max 2 times. Then use your own knowledge.
- When moving slides, plan ALL moves in ONE turn. Each slide can only be moved once.
- You can call multiple tools per turn.
- Colors: blue=#4A90D9, red=#E74C3C, green=#2ECC71, yellow=#F1C40F, orange=#E67E22, purple=#9B59B6, pink=#E91E63, black=#333333, white=#FFFFFF, dark navy=#1a1a2e.
- When moving slides, plan ALL moves in ONE turn. Each slide can only be moved once.
- DOCUMENT GENERATION: Slide is 960x540pt (widescreen 16:9). Use add_shape("Rectangle") for header/footer banners. Use add_text_box with precise left/top/width/height for structured layouts. Typical layout: header banner at top=0 (960x50), footer banner at top=500 (960x40), body text at left=60, top=70, width=840, height=410. For dense legal text, fontSize=8 or 9. For normal body, fontSize=10 or 11. For titles, fontSize=14-18 with bold.
- CRITICAL — CONTENT REQUIRED: When asked to create a "document", "disclaimer", "disclosure", "legal notice", or "report" page, you MUST generate the actual text content and put it in add_text_box or add_rich_text calls. Never create empty slides with just a title. Write real substantive text from your own knowledge. For legal disclaimers, write full multi-paragraph legal text. Each slide should have at least one body text box.
- RICH TEXT: Use add_rich_text for structured content where headings (bold, 14-18pt) and body paragraphs (normal, 10-11pt) appear together. Example: [{text:"一、版权声明",fontSize:16,bold:true},{text:"本文件...",fontSize:10}]. Use add_text_box for single-style text only.
- NEVER create duplicate slides with the same title. One slide = one title. You can only call add_slide/add_slide_with_title ONCE per unique title per turn.
- CALL add_slide_with_title ONLY ONCE. Then immediately populate that slide with shapes and text boxes. Do NOT call add_slide multiple times.
- MULTI-PAGE: Only create slides you have actual content for. Don't pre-create placeholder slides hoping to fill them later. If the text fits on 1-2 slides, only create 1-2 slides. If you need more slides, create a new slide AND immediately fill it with add_text_box in the same turn.
- When the task is fully complete (slides with actual content created), say so in a final message.`;

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
    messages.push({ role: "system", content: `Current: ${await buildSlideContext()}Use your own knowledge. Multi-slide: always use target_slide.` });
  }
  messages.push(...conversationHistory);
  messages.push({ role: "user", content: userCommand });

  for (let turn = 0; turn < MAX; turn++) {
    const resp = await fetch(`${API_BASE}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
      body: JSON.stringify({ model: MODEL, messages, tools: TOOLS, tool_choice: "auto", temperature: 0.1, max_tokens: 2000 }),
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

      // ── Dedup: block duplicate add_slide BEFORE dryRun check ──
      const isSlideAdd = tc.function.name === "add_slide" || tc.function.name === "add_slide_with_title";
      if (isSlideAdd && addedSlideTitles.has(args.title || "")) {
        result = { success: false, message: `⚠️ Duplicate slide "${args.title || "untitled"}" skipped.` };
      } else {
        // Track as used (even in dryRun, to prevent duplicates in planned actions)
        if (isSlideAdd) addedSlideTitles.add(args.title || "");

        // In dryRun mode, still execute read-only tools so AI gets real data
        if (dryRun && !isReadOnly) {
          result = { success: true, message: `[Preview] Would call ${tc.function.name}` };
        } else if (tc.function.name === "web_search" && searchCount >= 2) {
          result = { success: false, message: "⚠️ Maximum 2 web searches reached. Create slides NOW." };
        } else if (tc.function.name === "move_slide" && movedSlides.has(String(args.fromIndex))) {
          result = { success: false, message: `⚠️ Slide ${args.fromIndex} was already moved — cannot move again.` };
        } else {
          if (tc.function.name === "web_search") searchCount++;
          if (tc.function.name === "move_slide") movedSlides.add(String(args.fromIndex));
          result = await executeToolCall(tc.function.name, args, currentSlideId);
        }
      }

      // Only push to pendingCalls if not a duplicate
      if (result.success || !result.message.includes("skipped")) {
        toolResults.push(result);
        pendingCalls.push({ name: tc.function.name, args, description: `${tc.function.name}(${JSON.stringify(args).slice(0, 80)})` });
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

    const r = await executeToolCall(call.name, call.args, targetSid);
    results.push(r);

    // After adding a slide, auto-track it as default for subsequent tools
    if (r.success && (call.name === "add_slide" || call.name === "add_slide_with_title")) {
      try {
        const slides = await getSlides();
        if (slides.length > slideCount) {
          defaultSid = slides[slides.length - 1].id;
          slideCount = slides.length;
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
        await addShape(args.geometry, { fillColor: args.fillColor ? toHex(args.fillColor) : "#4A90D9", left: args.left, top: args.top, width: args.width, height: args.height });
        return { success: true, message: `Added ${args.geometry}` };

      case "add_image":
        await addImage(args.url, { left: args.left, top: args.top, width: args.width, height: args.height });
        return { success: true, message: `Added image from ${args.url}` };

      case "add_text_box":
        await addTextBox(args.text, { left: args.left, top: args.top, width: args.width, height: args.height, fontSize: args.fontSize });
        return { success: true, message: `Added text: "${args.text.slice(0, 60)}${args.text.length > 60 ? "..." : ""}"` };

      case "add_rich_text":
        await addStructuredTextBox(args.paragraphs, { left: args.left, top: args.top, width: args.width, height: args.height });
        return { success: true, message: `Added rich text: ${args.paragraphs.length} paragraph(s)` };

      case "modify_all_shapes": {
        if (!sid) return { success: false, message: "No slide selected" };
        const n = await applyStyleToAllShapes(sid, { fillColor: args.fillColor ? toHex(args.fillColor) : undefined, fontSize: args.fontSize });
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
        var tr = await upsertTable({ headers: args.headers, rows: args.rows });
        return { success: true, message: tr };

      case "add_chart":
        await addChart(args.chartType as ChartType, { categories: args.categories, series: args.series }, { title: args.title });
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

      case "no_op":
        return { success: false, message: args.message || "Cannot fulfill" };

      default:
        return { success: false, message: `Unknown tool: ${name}` };
    }
  } catch (err: any) {
    return { success: false, message: `${name}: ${err.message || err}` };
  }
}
