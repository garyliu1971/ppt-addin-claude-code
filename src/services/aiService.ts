/**
 * aiService.ts — AI-powered command processing
 * Multi-turn conversation with DeepSeek + web search via DuckDuckGo.
 */

import {
  getSlides, getSelectedSlide, getShapesOnSlide, getPresentationInfo,
} from "./pptApi";
import {
  addShape, addTextBox, setShapeFill, deleteShape, applyStyleToAllShapes,
} from "./shapeService";
import { addTable, addChart, ChartType, ChartData } from "./chartTableService";
import {
  applyLayoutToSlide, findLayoutByName, setSlideBackground,
  addSlide, deleteSlide, deleteSlideByIndex, setSlideTitle,
  moveSlide, duplicateSlide, addSlideWithTitle, getSlidesWithIndex,
} from "./masterLayoutThemeService";

// ── Config ────────────────────────────────────────────────────────

const STORAGE_KEY = "ppt_ai_api_key";
let apiKey = "";
try { const s = localStorage.getItem(STORAGE_KEY); if (s) apiKey = s; } catch { /* */ }

const API_BASE = "https://api.deepseek.com/v1";
const MODEL = "deepseek-chat";

export function setApiKey(k: string): void { apiKey = k; try { localStorage.setItem(STORAGE_KEY, k); } catch { /* */ } }
export function clearApiKey(): void { apiKey = ""; try { localStorage.removeItem(STORAGE_KEY); } catch { /* */ } }
export function getApiKey(): string { return apiKey; }
export function hasApiKey(): boolean { return apiKey.length > 0; }

// ── Web Search (multi-source: CoinGecko, Wikipedia, DuckDuckGo) ─────

async function searchWeb(query: string): Promise<string> {
  const q = query.toLowerCase();

  // ── Crypto prices: CoinGecko (free, CORS-enabled) ──────────────
  if (/bitcoin|btc|crypto|ethereum|eth|doge|solana|price|价格/i.test(q)) {
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

  return "Search unavailable. Try a more specific query.";
}

// ── Tool Definitions ──────────────────────────────────────────────

interface ToolDef {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, any> };
}

const TOOLS: ToolDef[] = [
  { type: "function", function: { name: "add_shape", description: "Add a geometric shape. Types: Rectangle, Oval, Triangle, Diamond, Arrow, Star5, Heart, Cloud, Sun, Moon, SmileyFace.", parameters: { type: "object", properties: { geometry: { type: "string" }, fillColor: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" } }, required: ["geometry"] } } },
  { type: "function", function: { name: "add_text_box", description: "Add a text box to the current slide.", parameters: { type: "object", properties: { text: { type: "string" }, left: { type: "number" }, top: { type: "number" } }, required: ["text"] } } },
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
  { type: "function", function: { name: "web_search", description: "Search the web for real-time data (sports scores, weather, news, stock prices). Use when you need current information.", parameters: { type: "object", properties: { query: { type: "string" } }, required: ["query"] } } },
  { type: "function", function: { name: "no_op", description: "Use when request cannot be fulfilled.", parameters: { type: "object", properties: { message: { type: "string" } }, required: ["message"] } } },
];

// ── Context Builder ───────────────────────────────────────────────

async function buildSlideContext(): Promise<string> {
  try {
    const info = await getPresentationInfo();
    const slide = await getSelectedSlide();
    let ctx = `Presentation: "${info.title}" | ${info.slideCount} slide(s)\n`;
    if (slide) {
      ctx += `Current slide: ${slide.id}\n`;
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

export interface AIResult {
  messages: string[];
  toolResults: ExecResult[];
}

export async function runAIConversation(
  userCommand: string,
  currentSlideId: string | null
): Promise<AIResult> {
  if (!apiKey) throw new Error("API key not set.");

  const sysPrompt = `You are a PowerPoint AI assistant. Use tools to fulfill requests.
Current context: ${await buildSlideContext()}

Rules:
- For real-time data (sports, weather, news, stocks), use web_search FIRST.
- You can call multiple tools. After each tool you'll see the result.
- Colors: blue=#4A90D9, red=#E74C3C, green=#2ECC71, yellow=#F1C40F, orange=#E67E22, purple=#9B59B6, pink=#E91E63, black=#333333, white=#FFFFFF.
- When done, give a final summary. Do not call tools in the final message.`;

  const messages: any[] = [
    { role: "system", content: sysPrompt },
    { role: "user", content: userCommand },
  ];

  const textMessages: string[] = [];
  const toolResults: ExecResult[] = [];
  const MAX = 6;

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
      const result = await executeToolCall(tc.function.name, args, currentSlideId);
      toolResults.push(result);
      messages.push({
        role: "tool", tool_call_id: tc.id,
        content: JSON.stringify({ success: result.success, message: result.message, context: await buildSlideContext() }),
      });
    }
  }

  return { messages: textMessages, toolResults };
}

// ── Tool Executor ─────────────────────────────────────────────────

const CM: Record<string, string> = { blue: "#4A90D9", red: "#E74C3C", green: "#2ECC71", yellow: "#F1C40F", orange: "#E67E22", purple: "#9B59B6", pink: "#E91E63", black: "#333333", white: "#FFFFFF", gray: "#95A5A6" };
const toHex = (c: string) => CM[c.toLowerCase()] || c;

export async function executeToolCall(
  name: string, args: Record<string, any>, sid: string | null
): Promise<ExecResult> {
  try {
    switch (name) {
      case "add_shape":
        await addShape(args.geometry, { fillColor: args.fillColor ? toHex(args.fillColor) : "#4A90D9", left: args.left, top: args.top, width: args.width, height: args.height });
        return { success: true, message: `Added ${args.geometry}` };

      case "add_text_box":
        await addTextBox(args.text, { left: args.left, top: args.top });
        return { success: true, message: `Added text: "${args.text}"` };

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
        await addTable({ headers: args.headers, rows: args.rows });
        return { success: true, message: `Added ${args.rows.length + 1}x${args.headers.length} table` };

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

      case "no_op":
        return { success: false, message: args.message || "Cannot fulfill" };

      default:
        return { success: false, message: `Unknown tool: ${name}` };
    }
  } catch (err: any) {
    return { success: false, message: `${name}: ${err.message || err}` };
  }
}
