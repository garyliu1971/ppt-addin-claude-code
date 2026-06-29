/**
 * taskpane.ts — Main Claude Code for PowerPoint task pane
 *
 * Provides:
 * - Natural language command parsing (Claude Code style)
 * - Quick action buttons
 * - Live event log from shape/slide changes
 * - Integration with all service modules
 */

import "./taskpane.css";
import { LogEntry, ensureOfficeReady, getSlides, getSelectedSlide, getShapesOnSlide, getPresentationInfo, getPageSetup } from "./services/pptApi";
import { addShape, addTextBox, setShapeFill, setShapeText, deleteShape, applyStyleToAllShapes, setShapeGeometry, setShapeFontSize, autoLayoutShapes } from "./services/shapeService";
import { addTable, addChart, ChartType, ChartData } from "./services/chartTableService";
import { getMasterDetails, applyLayoutToSlide, findLayoutByName, setSlideBackground, getThemeDetails, addSlide, deleteSlide, getAllLayouts, deleteSlideByIndex, setSlideTitle, moveSlide, duplicateSlide, addSlideWithTitle, getSlidesWithIndex, applyTheme, listAvailableThemes, applyDesignScheme, listDesignSchemes } from "./services/masterLayoutThemeService";
import { registerEventHandlers, unregisterEventHandlers } from "./services/eventService";
import { setApiKey, getApiKey, hasApiKey, runAIConversation, executePendingCalls, clearApiKey, clearConversationHistory, getHistoryLength, PendingCall } from "./services/aiService";

// ── State ─────────────────────────────────────────────────────────

let currentSlideId: string | null = null;
let aiEnabled: boolean = false;
let pendingActions: { name: string; args: any; description: string }[] = [];

// ── DOM Elements ──────────────────────────────────────────────────

const commandInput = document.getElementById("command-input") as HTMLTextAreaElement;
const sendBtn = document.getElementById("send-btn") as HTMLButtonElement;
const commandHistory = document.getElementById("command-history") as HTMLDivElement;
const outputContent = document.getElementById("output-content") as HTMLPreElement;
const statusIndicator = document.getElementById("status-indicator") as HTMLSpanElement;
const aiBadge = document.getElementById("ai-badge") as HTMLSpanElement;
const actionButtons = document.querySelectorAll<HTMLButtonElement>(".action-btn");
const apiKeyToggle = document.getElementById("api-key-toggle") as HTMLDivElement;
const apiKeyBody = document.getElementById("api-key-body") as HTMLDivElement;
const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
const apiKeySave = document.getElementById("api-key-save") as HTMLButtonElement;
const aiStatusText = document.getElementById("ai-status-text") as HTMLSpanElement;
const apiKeyStatus = document.getElementById("api-key-status") as HTMLSpanElement;

// ── Logging ───────────────────────────────────────────────────────

function log(entry: LogEntry): void {
  const timestamp = new Date(entry.timestamp).toLocaleTimeString();
  const prefix = {
    info: "ℹ️",
    warn: "⚠️",
    error: "❌",
    success: "✅",
  }[entry.level];

  const line = `[${timestamp}] ${prefix} ${entry.message}`;
  outputContent.textContent += "\n" + line;
  outputContent.scrollTop = outputContent.scrollHeight;

  // Also add to command history
  addHistoryEntry(entry);
}

function addHistoryEntry(entry: LogEntry): void {
  // command-history div removed in new layout — log goes to output-content only
  if (!commandHistory) return;
  const div = document.createElement("div");
  div.className = `command-entry ${entry.level === "error" ? "error" : "response"}`;
  div.textContent = entry.message;
  commandHistory.appendChild(div);
  commandHistory.scrollTop = commandHistory.scrollHeight;

  // Keep max 50 entries
  while (commandHistory.children.length > 50) {
    commandHistory.removeChild(commandHistory.firstChild!);
  }
}

// ── Command Dispatcher ────────────────────────────────────────────

/** Main command entry — tries AI first, falls back to regex */
async function executeCommand(input: string): Promise<void> {
  const cmd = input.trim();
  if (!cmd) return;

  // Always refresh current slide before executing
  try {
    const active = await getSelectedSlide();
    if (active) currentSlideId = active.id;
  } catch { /* ignore */ }

  // Add user command to history (skip if command-history removed)
  if (commandHistory) {
    const userDiv = document.createElement("div");
    userDiv.className = "command-entry user";
    userDiv.textContent = "🧠 " + cmd;
    commandHistory.appendChild(userDiv);
    commandHistory.scrollTop = commandHistory.scrollHeight;
  }

  // ── AI MODE ────────────────────────────────────────────────────
  if (aiEnabled && hasApiKey()) {
    log({ level: "info", message: `🤖 AI thinking: "${cmd}"`, timestamp: Date.now() });
    sendBtn.disabled = true;
    sendBtn.textContent = "⏳";

    try {
      // Step 1: Dry run — AI plans but doesn't execute
      const result = await runAIConversation(cmd, currentSlideId, true);

      // Show AI text responses
      for (const text of result.messages) {
        log({ level: "info", message: text, timestamp: Date.now() });
      }

      // Show planned actions summary
      if (result.pendingCalls.length > 0) {
        const actions = result.pendingCalls
          .filter((c: PendingCall) => c.name !== "list_slides" && c.name !== "list_themes" && c.name !== "no_op" && c.name !== "web_search")
          .map((c: PendingCall) => `  ⏳ ${c.description}`);

        const searchCalls = result.pendingCalls.filter((c: PendingCall) => c.name === "web_search");
        if (searchCalls.length > 0) {
          log({ level: "info", message: `🔍 ${searchCalls.length} web search(es) performed`, timestamp: Date.now() });
        }

        if (actions.length > 0) {
          log({ level: "info", message: `📋 Planned actions:\n${actions.join("\n")}`, timestamp: Date.now() });
          // Store pending calls and show confirm bar
          pendingActions = result.pendingCalls;
          showConfirmBar(result.pendingCalls.length);
        } else {
          log({ level: "info", message: "✅ No slide changes needed.", timestamp: Date.now() });
          sendBtn.disabled = false;
          sendBtn.textContent = "▶ Send";
        }
      } else {
        log({ level: "info", message: "✅ Done.", timestamp: Date.now() });
        sendBtn.disabled = false;
        sendBtn.textContent = "▶ Send";
      }
    } catch (err: any) {
      log({ level: "error", message: `AI error: ${err.message || err}`, timestamp: Date.now() });
      sendBtn.disabled = false;
      sendBtn.textContent = "▶ Send";
    }
    return;
  }

  // ── REGEX FALLBACK ─────────────────────────────────────────────
  if (!aiEnabled && hasApiKey()) {
    log({ level: "warn", message: '💡 AI key is saved but AI mode is OFF. Click 🧠 in header to enable AI.', timestamp: Date.now() });
  }
  log({ level: "info", message: aiEnabled ? "⚠️ AI not configured — using regex" : `Executing: "${cmd}"`, timestamp: Date.now() });

  try {
    const lower = cmd.toLowerCase();

    // ── Shape Commands ────────────────────────────────────────────

    if (/make all shapes? (.*)/i.test(cmd) || /set all shapes? (.*)/i.test(cmd)) {
      // "make all shapes blue" / "set all shapes to red"
      const colorMatch = cmd.match(/(blue|red|green|yellow|orange|purple|pink|black|white|gray|grey)/i);
      if (colorMatch && currentSlideId) {
        const colorMap: Record<string, string> = {
          blue: "#4A90D9", red: "#E74C3C", green: "#2ECC71", yellow: "#F1C40F",
          orange: "#E67E22", purple: "#9B59B6", pink: "#E91E63", black: "#333333",
          white: "#FFFFFF", gray: "#95A5A6", grey: "#95A5A6",
        };
        const color = colorMap[colorMatch[1].toLowerCase()] || "#4A90D9";
        const count = await applyStyleToAllShapes(currentSlideId, { fillColor: color });
        log({ level: "success", message: `Applied ${colorMatch[1]} fill to ${count} shape(s)`, timestamp: Date.now() });
      } else {
        log({ level: "warn", message: "Select a slide first, then specify a color.", timestamp: Date.now() });
      }
    }
    else if (/add (a |an )?(rectangle|oval|triangle|diamond|arrow|heart|star5|circle|square)( shape)?/i.test(cmd)) {
      const shapeMatch = cmd.match(/add (a |an )?(rectangle|oval|triangle|diamond|arrow|heart|star5|circle|square)( shape)?/i);
      const shapeType = shapeMatch![2];
      const colorMatch = cmd.match(/(blue|red|green|yellow|orange|purple|pink)/i);

      try {
        // Capitalize first letter for enum
        const geometry = (shapeType.charAt(0).toUpperCase() + shapeType.slice(1)) as any;
        const shape = await addShape(geometry, {
          fillColor: colorMatch ? getColorHex(colorMatch[1]) : "#4A90D9",
        });
        log({ level: "success", message: `Added ${shapeType} shape (id: ${shape.id})`, timestamp: Date.now() });
      } catch {
        log({ level: "error", message: `Unknown shape type: "${shapeType}". Try: rectangle, oval, triangle, diamond, arrow, heart, star5`, timestamp: Date.now() });
      }
    }
    else if (/add (a |an )?text\s?box (.+)/i.test(cmd)) {
      const textMatch = cmd.match(/add (a |an )?text\s?box ["']?(.+?)["']?$/i);
      const text = textMatch![2] || cmd.replace(/add (a |an )?text\s?box/i, "").trim();
      const shape = await addTextBox(text);
      log({ level: "success", message: `Added text box: "${text}" (id: ${shape.id})`, timestamp: Date.now() });
    }
    else if (/delete (the )?selected shape/i.test(cmd)) {
      if (currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length > 0) {
          await deleteShape(shapes[shapes.length - 1].id, currentSlideId);
          log({ level: "success", message: "Deleted last shape on slide", timestamp: Date.now() });
        }
      }
    }
    else if (/fill shape (.+) with (.+)/i.test(cmd)) {
      const fillMatch = cmd.match(/fill shape ["']?(.+?)["']? with (.+)/i);
      if (fillMatch && currentSlideId) {
        const shapeName = fillMatch[1];
        const colorName = fillMatch[2].trim();
        const shapes = await getShapesOnSlide(currentSlideId);
        const target = shapes.find((s) => s.name?.toLowerCase().includes(shapeName.toLowerCase()));
        if (target) {
          await setShapeFill(target.id, currentSlideId, getColorHex(colorName));
          log({ level: "success", message: `Filled "${target.name}" with ${colorName}`, timestamp: Date.now() });
        } else {
          log({ level: "warn", message: `Shape "${shapeName}" not found`, timestamp: Date.now() });
        }
      }
    }
    else if (/resize shape/i.test(cmd)) {
      const sizeMatch = cmd.match(/(\d+)\s*[x×*X]\s*(\d+)/);
      if (sizeMatch && currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        // Resize the most recently added shape
        const target = shapes[shapes.length - 1];
        if (target) {
          await setShapeGeometry(target.id, currentSlideId, {
            width: parseInt(sizeMatch[1]),
            height: parseInt(sizeMatch[2]),
          });
          log({ level: "success", message: `Resized shape to ${sizeMatch[1]}x${sizeMatch[2]}`, timestamp: Date.now() });
        } else {
          log({ level: "warn", message: "No shapes on current slide.", timestamp: Date.now() });
        }
      } else if (!sizeMatch) {
        log({ level: "warn", message: "Usage: resize shape WxH  (e.g. resize shape 200x150)", timestamp: Date.now() });
      }
    }
    else if (/(set )?(font|text) size\s*(\d+)/i.test(cmd)) {
      const fsMatch = cmd.match(/(?:set )?(?:font|text) size\s*(\d+)/i);
      const fontSize = parseInt(fsMatch![1]);
      if (currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length === 0) {
          log({ level: "warn", message: "No shapes on current slide.", timestamp: Date.now() });
        } else {
          const target = shapes[shapes.length - 1];
          await setShapeFontSize(target.id, currentSlideId, fontSize);
          log({ level: "success", message: `Set font size to ${fontSize}pt on "${target.name || "unnamed"}"`, timestamp: Date.now() });
        }
      } else {
        log({ level: "warn", message: "Select a slide first.", timestamp: Date.now() });
      }
    }
    else if (/bold\s*text/i.test(cmd)) {
      if (currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length === 0) {
          log({ level: "warn", message: "No shapes on current slide.", timestamp: Date.now() });
        } else {
          const target = shapes[shapes.length - 1];
          await setShapeText(target.id, currentSlideId, "", { bold: true });
          log({ level: "success", message: `Set bold on "${target.name || "unnamed"}"`, timestamp: Date.now() });
        }
      }
    }
    else if (/set (text|font) color/i.test(cmd)) {
      const colorMatch = cmd.match(/(blue|red|green|yellow|orange|purple|pink|black|white|gray)/i);
      if (colorMatch && currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length > 0) {
          const target = shapes[shapes.length - 1];
          await setShapeText(target.id, currentSlideId, "", { fontColor: getColorHex(colorMatch[1]) });
          log({ level: "success", message: `Set text color to ${colorMatch[1]}`, timestamp: Date.now() });
        }
      }
    }

    // ── Table Commands ────────────────────────────────────────────

    else if (/add (a |an )?table/i.test(cmd)) {
      await addTable(
        {
          headers: ["Item", "Value", "Notes"],
          rows: [
            ["Row 1", "100", "Sample"],
            ["Row 2", "200", "Sample"],
            ["Row 3", "300", "Sample"],
          ],
        }
      );
      log({ level: "success", message: "Added 3×4 table to slide", timestamp: Date.now() });
    }

    // ── Chart Commands ────────────────────────────────────────────

    else if (/add (a |an )?(bar|column|pie|line|doughnut|area) chart/i.test(cmd)) {
      const chartMatch = cmd.match(/add (a |an )?(bar|column|pie|line|doughnut|area) chart/i);
      const type = chartMatch![2];
      const chartTypeMap: Record<string, ChartType> = {
        bar: "BarClustered", column: "ColumnClustered", pie: "Pie",
        line: "Line", doughnut: "Doughnut", area: "Area",
      };

      const chartData: ChartData = {
        categories: ["Q1", "Q2", "Q3", "Q4"],
        series: [
          { name: "Revenue", values: [1200, 1500, 1800, 2100] },
          { name: "Cost", values: [800, 900, 1000, 1100] },
        ],
      };

      await addChart(chartTypeMap[type] || "ColumnClustered", chartData, {
        title: `${type.charAt(0).toUpperCase() + type.slice(1)} Chart`,
      });
      log({ level: "success", message: `Added ${type} chart to slide`, timestamp: Date.now() });
    }

    // ── Slide / Layout Commands ───────────────────────────────────

    else if (/apply layout/i.test(cmd)) {
      if (currentSlideId) {
        const layoutMatch = cmd.match(/layout ["']?(.+?)["']?$/i);
        if (layoutMatch) {
          const found = await findLayoutByName(layoutMatch[1]);
          if (found) {
            await applyLayoutToSlide(currentSlideId, found.layoutId);
            log({ level: "success", message: `Applied layout "${layoutMatch[1]}"`, timestamp: Date.now() });
          } else {
            log({ level: "warn", message: `Layout "${layoutMatch[1]}" not found`, timestamp: Date.now() });
          }
        }
      }
    }
    else if (/set (slide )?background/i.test(cmd)) {
      if (currentSlideId) {
        const colorMatch = cmd.match(/(blue|red|green|yellow|orange|purple|pink|black|white|gray)/i);
        if (colorMatch) {
          await setSlideBackground(currentSlideId, getColorHex(colorMatch[1]));
          log({ level: "success", message: `Set slide background to ${colorMatch[1]}`, timestamp: Date.now() });
        }
      }
    }
    else if (/add (a |an )?slide/i.test(cmd)) {
      // "add slide" or "add slide titled My Title"
      const titleMatch = cmd.match(/(?:slide\s+(?:titled?|named?|with\s+title)\s+)[\"']?(.+?)[\"']?$/i);
      if (titleMatch) {
        const s = await addSlideWithTitle(titleMatch[1]);
        log({ level: "success", message: `Added slide: "${titleMatch[1]}" (id: ${s.id})`, timestamp: Date.now() });
      } else {
        const newSlide = await addSlide();
        log({ level: "success", message: `Added new slide (id: ${newSlide.id})`, timestamp: Date.now() });
      }
    }
    else if (/delete slide\s*(\d+)/i.test(cmd)) {
      // "delete slide 3"
      const numMatch = cmd.match(/delete slide\s*(\d+)/i);
      const idx = parseInt(numMatch![1]);
      const msg = await deleteSlideByIndex(idx);
      log({ level: "success", message: msg, timestamp: Date.now() });
    }
    else if (/delete (this |current )?slide/i.test(cmd)) {
      if (currentSlideId) {
        await deleteSlide(currentSlideId);
        currentSlideId = null;
        log({ level: "success", message: "Deleted current slide", timestamp: Date.now() });
      } else {
        log({ level: "warn", message: "No slide selected.", timestamp: Date.now() });
      }
    }
    else if (/set (slide )?title/i.test(cmd)) {
      if (!currentSlideId) { log({ level: "warn", message: "No slide selected.", timestamp: Date.now() }); }
      else {
        const titleMatch = cmd.match(/(?:title\s+(?:to\s+)?)[\"']?(.+?)[\"']?$/i);
        if (titleMatch) {
          await setSlideTitle(currentSlideId, titleMatch[1]);
          log({ level: "success", message: `Slide title set to: "${titleMatch[1]}"`, timestamp: Date.now() });
        } else {
          log({ level: "warn", message: 'Usage: "set title to Your Title Here"', timestamp: Date.now() });
        }
      }
    }
    else if (/move slide\s*(\d+)\s+(?:to\s+)?(\d+)/i.test(cmd)) {
      const mvMatch = cmd.match(/move slide\s*(\d+)\s+(?:to\s+)?(\d+)/i);
      const slides = await getSlidesWithIndex();
      const from = slides.find(s => s.index === parseInt(mvMatch![1]));
      if (!from) { log({ level: "warn", message: `Slide ${mvMatch![1]} not found`, timestamp: Date.now() }); }
      else {
        await moveSlide(from.id, parseInt(mvMatch![2]));
        log({ level: "success", message: `Moved slide ${mvMatch![1]} → ${mvMatch![2]}`, timestamp: Date.now() });
      }
    }
    else if (/duplicate (this |current )?slide/i.test(cmd)) {
      if (!currentSlideId) { log({ level: "warn", message: "No slide selected.", timestamp: Date.now() }); }
      else {
        await duplicateSlide(currentSlideId);
        log({ level: "success", message: "Duplicated current slide", timestamp: Date.now() });
      }
    }
    else if (/应用主题\s+(.+)|apply theme\s+(.+)/i.test(cmd)) {
      const tMatch = cmd.match(/(?:应用主题|apply theme)\s+(.+)/i);
      if (tMatch) {
        const msg = await applyTheme(tMatch[1].trim());
        log({ level: "success", message: msg, timestamp: Date.now() });
      }
    }
    else if (/应用设计\s+(.+)|design\s+(.+)|设计\s+(.+)/i.test(cmd)) {
      if (currentSlideId) {
        const dMatch = cmd.match(/(?:应用设计|design|设计)\s+(.+)/i);
        if (dMatch) {
          const msg = await applyDesignScheme(currentSlideId, dMatch[1].trim());
          log({ level: "success", message: msg, timestamp: Date.now() });
        }
      } else {
        log({ level: "warn", message: "请先选择一个幻灯片", timestamp: Date.now() });
      }
    }
    else if (/list themes|列出主题|主题列表/i.test(cmd)) {
      const themes = listAvailableThemes();
      const schemes = listDesignSchemes();
      log({ level: "info", message: `🎨 主题 (${themes.length}): ${themes.join(", ")}\n🎨 设计方案 (${schemes.length}): ${schemes.join(", ")}`, timestamp: Date.now() });
    }

    // ── Info / Help ───────────────────────────────────────────────

    else if (/help/i.test(cmd)) {
      log({ level: "info", message: `Available commands:
  SHAPES:
• add [rectangle|oval|triangle|diamond|arrow|heart|star5|circle|square] [shape] [color]
• add text box "your text"
• make all shapes [color]
• fill shape "name" with [color]
• resize shape W×H
• delete shape "name"  TEXT:
• font size N              (e.g. "font size 10")
• bold text
• set text color [color]  CHARTS & TABLES:
• add [bar|column|pie|line|area] chart
• add table
  SLIDES:
• add slide [titled "Title"]
• delete slide [N]      (e.g. "delete slide 3")
• delete current slide
• set title to "New Title"
• move slide N to M     (e.g. "move slide 1 to 3")
• duplicate slide
• apply layout "name"
• set background [color]
  INFO:
• list slides / shapes / layouts / masters
• page setup / slide size
• show theme`, timestamp: Date.now() });
    }
    else if (/list (all )?slides?/i.test(cmd)) {
      const slides = await getSlides();
      const info = slides.map((s, i) => `  Slide ${i + 1}: id=${s.id}`).join("\n");
      log({ level: "info", message: `${slides.length} slide(s):\n${info}`, timestamp: Date.now() });
    }
    else if (/list (all )?shapes?/i.test(cmd)) {
      if (currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        const list = shapes.map((s) => `  ${s.name || "unnamed"} (${s.type})`).join("\n");
        log({ level: "info", message: `${shapes.length} shape(s):\n${list}`, timestamp: Date.now() });
      } else {
        log({ level: "warn", message: "Select a slide first.", timestamp: Date.now() });
      }
    }
    else if (/list (all )?layouts?/i.test(cmd)) {
      const layouts = await getAllLayouts();
      const list = layouts.map((l) => `  [${l.masterName}] ${l.layoutName}`).join("\n");
      log({ level: "info", message: `${layouts.length} layout(s):\n${list}`, timestamp: Date.now() });
    }
    else if (/list (all )?masters?/i.test(cmd)) {
      const masters = await getMasterDetails();
      const list = masters.map((m) => `  ${m.name} (${m.layoutCount} layouts)`).join("\n");
      log({ level: "info", message: `${masters.length} master(s):\n${list}`, timestamp: Date.now() });
    }
    else if (/(page|slide) (setup|size)|页面设置|幻灯片尺寸/i.test(cmd)) {
      const ps = await getPageSetup();
      if (ps) {
        const inchesW = (ps.width / 72).toFixed(1);
        const inchesH = (ps.height / 72).toFixed(1);
        log({ level: "info", message: `📐 Slide size: ${ps.width}×${ps.height} pt (${inchesW}×${inchesH} in)`, timestamp: Date.now() });
      } else {
        log({ level: "warn", message: "Page setup info not available (requires Office 365).", timestamp: Date.now() });
      }
    }

    // ── Chinese Commands (中文命令) ────────────────────────────────

    else if (/删[除掉]?\s*(第\s*(\d+|[一二三四五六七八九十]+)\s*页|当前页|这页)/.test(cmd)) {
      const numMatch = cmd.match(/第\s*(\d+|[一二三四五六七八九十]+)\s*页/);
      if (numMatch) {
        const cnNums: Record<string, number> = { 一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10 };
        const idx = cnNums[numMatch[1]] ?? parseInt(numMatch[1]);
        const msg = await deleteSlideByIndex(idx);
        log({ level: "success", message: msg, timestamp: Date.now() });
      } else if (currentSlideId) {
        await deleteSlide(currentSlideId);
        currentSlideId = null;
        log({ level: "success", message: "已删除当前页", timestamp: Date.now() });
      }
    }
    else if (/(添加|新建|新增|插入)\s*(一张|一个)?\s*幻灯[片页]/i.test(cmd)) {
      const titleMatch = cmd.match(/(?:标题[是为]?\s*)[\"']?(.+?)[\"']?$/);
      if (titleMatch) {
        const s = await addSlideWithTitle(titleMatch[1]);
        log({ level: "success", message: `已添加幻灯片: "${titleMatch[1]}"`, timestamp: Date.now() });
      } else {
        await addSlide();
        log({ level: "success", message: "已添加新幻灯片", timestamp: Date.now() });
      }
    }
    else if (/移动\s*第?\s*(\d+|[一二三四五六七八九十]+)\s*(页|个)?\s*(到|至)\s*第?\s*(\d+|[一二三四五六七八九十]+)/i.test(cmd)) {
      const mvMatch = cmd.match(/移动\s*第?\s*(\d+|[一二三四五六七八九十]+)\s*(?:页|个)?\s*(?:到|至)\s*第?\s*(\d+|[一二三四五六七八九十]+)/i);
      if (mvMatch) {
        const cnNums: Record<string, number> = { 一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10 };
        const from = cnNums[mvMatch[1]] ?? parseInt(mvMatch[1]);
        const to = cnNums[mvMatch[2]] ?? parseInt(mvMatch[2]);
        const slides = await getSlidesWithIndex();
        const src = slides.find(s => s.index === from);
        if (src) {
          await moveSlide(src.id, to);
          log({ level: "success", message: `已将第${from}页移至第${to}页`, timestamp: Date.now() });
        }
      }
    }
    else if (/(复制|拷贝|克隆)\s*(当前|这[张个])?\s*幻灯[片页]/i.test(cmd)) {
      if (currentSlideId) {
        await duplicateSlide(currentSlideId);
        log({ level: "success", message: "已复制当前幻灯片", timestamp: Date.now() });
      }
    }
    else if (/(设置|修改|更改|改)\s*(幻灯[片页]?\s*)?标题/i.test(cmd)) {
      if (!currentSlideId) { log({ level: "warn", message: "请先选择一个幻灯片", timestamp: Date.now() }); }
      else {
        const titleMatch = cmd.match(/(?:标题[是为]?\s*)[\"']?(.+?)[\"']?$/);
        if (titleMatch) {
          await setSlideTitle(currentSlideId, titleMatch[1]);
          log({ level: "success", message: `标题已设置为: "${titleMatch[1]}"`, timestamp: Date.now() });
        } else {
          log({ level: "warn", message: '用法: "设置标题为 你的标题"', timestamp: Date.now() });
        }
      }
    }
    else if (/(设置|更改)\s*背景/i.test(cmd)) {
      if (currentSlideId) {
        const colorMatch = cmd.match(/(蓝|红|绿|黄|橙|紫|粉|黑|白|灰|blue|red|green|yellow|orange|purple|pink|black|white|gray)/i);
        const cnColorMap: Record<string, string> = { 蓝:"blue",红:"red",绿:"green",黄:"yellow",橙:"orange",紫:"purple",粉:"pink",黑:"black",白:"white",灰:"gray" };
        const color = colorMatch ? (cnColorMap[colorMatch[1]] || colorMatch[1]) : "black";
        await setSlideBackground(currentSlideId, getColorHex(color));
        log({ level: "success", message: `背景已设置为${color}`, timestamp: Date.now() });
      }
    }
    else if (/(列出|显示|查看)\s*(所有)?\s*幻灯[片页]/i.test(cmd)) {
      const slides = await getSlides();
      const info = slides.map((s, i) => `  第${i+1}页: id=${s.id}`).join("\n");
      log({ level: "info", message: `共 ${slides.length} 页:\n${info}`, timestamp: Date.now() });
    }
    else if (/(列出|显示|查看)\s*(所有)?\s*形状/i.test(cmd)) {
      if (currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        const list = shapes.map((s) => `  ${s.name || "未命名"} (${s.type})`).join("\n");
        log({ level: "info", message: `共 ${shapes.length} 个形状:\n${list}`, timestamp: Date.now() });
      }
    }
    else if (/(重排|重新排列|自动排列|auto.layout|整理)\s*(形状|所有)?/.test(cmd)) {
      if (currentSlideId) {
        const n = await autoLayoutShapes(currentSlideId);
        log({ level: "success", message: `已自动重排 ${n} 个形状`, timestamp: Date.now() });
      } else {
        log({ level: "warn", message: "请先选择一个幻灯片", timestamp: Date.now() });
      }
    }
    else if (/(字体大小|字号)\s*(\d+)/i.test(cmd)) {
      const fsMatch = cmd.match(/(?:字体大小|字号)\s*(\d+)/i);
      if (fsMatch && currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length > 0) {
          const fontSize = parseInt(fsMatch[1]);
          await setShapeFontSize(shapes[shapes.length - 1].id, currentSlideId, fontSize);
          log({ level: "success", message: `字体大小已设为 ${fontSize}pt`, timestamp: Date.now() });
        }
      }
    }
    else if (/粗体|加粗|bold/i.test(cmd)) {
      if (currentSlideId) {
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length > 0) {
          await setShapeText(shapes[shapes.length - 1].id, currentSlideId, "", { bold: true });
          log({ level: "success", message: "已设粗体", timestamp: Date.now() });
        }
      }
    }
    else if (/文字颜色\s*(蓝|红|绿|黄|橙|紫|粉|黑|白|灰)/i.test(cmd)) {
      const m = cmd.match(/文字颜色\s*(蓝|红|绿|黄|橙|紫|粉|黑|白|灰)/i);
      if (m && currentSlideId) {
        const cnColorMap: Record<string, string> = { 蓝:"blue",红:"red",绿:"green",黄:"yellow",橙:"orange",紫:"purple",粉:"pink",黑:"black",白:"white",灰:"gray" };
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length > 0) {
          await setShapeText(shapes[shapes.length - 1].id, currentSlideId, "", { fontColor: getColorHex(cnColorMap[m[1]] || m[1]) });
          log({ level: "success", message: `文字颜色已设为${m[1]}`, timestamp: Date.now() });
        }
      }
    }
    else if (/帮助|help/i.test(cmd)) {
      log({ level: "info", message: `📋 支持的命令:
  形状: 添加矩形|添加圆形|添加三角形|添加文本框 "内容"
        所有形状设为蓝色|填充形状 "名称" 红色
  文字: 字体大小 10|粗体|文字颜色 蓝色
  图表: 添加柱状图|添加饼图|添加表格
  幻灯片: 添加幻灯片 [标题为 "xxx"]|删掉第N页|删掉当前页
          移动第N页到第M页|复制当前页|设置标题为 "xxx"
          设置背景 蓝色|列出幻灯片|列出形状
          页面设置|幻灯片尺寸`, timestamp: Date.now() });
    }

    else {
      log({ level: "warn", message: `Unknown command: "${cmd}". Type "help" for available commands.`, timestamp: Date.now() });
    }
  } catch (err: any) {
    log({ level: "error", message: `Error: ${err.message || err}`, timestamp: Date.now() });
  }
}

function getColorHex(name: string): string {
  const map: Record<string, string> = {
    blue: "#4A90D9", red: "#E74C3C", green: "#2ECC71", yellow: "#F1C40F",
    orange: "#E67E22", purple: "#9B59B6", pink: "#E91E63", black: "#333333",
    white: "#FFFFFF", gray: "#95A5A6", grey: "#95A5A6",
  };
  return map[name.toLowerCase()] || "#4A90D9";
}

// ── Quick Action Handlers ─────────────────────────────────────────

async function handleQuickAction(action: string): Promise<void> {
  log({ level: "info", message: `Quick action: ${action}`, timestamp: Date.now() });

  // Always refresh current slide before executing
  try {
    const active = await getSelectedSlide();
    if (active) currentSlideId = active.id;
  } catch { /* ignore */ }

  try {
    switch (action) {
      case "listShapes": {
        if (currentSlideId) {
          const shapes = await getShapesOnSlide(currentSlideId);
          if (shapes.length === 0) {
            log({ level: "info", message: "No shapes on current slide.", timestamp: Date.now() });
          } else {
            for (const s of shapes) {
              log({ level: "info", message: `📐 ${s.name || "unnamed"} | type=${s.type} | pos=(${s.left},${s.top}) | size=${s.width}×${s.height}`, timestamp: Date.now() });
            }
          }
        } else {
          log({ level: "warn", message: "No slide selected.", timestamp: Date.now() });
        }
        break;
      }
      case "listSlides": {
        const slides = await getSlides();
        slides.forEach((s, i) => {
          log({ level: "info", message: `📑 Slide ${i + 1}: id=${s.id}`, timestamp: Date.now() });
        });
        break;
      }
      case "listMasters": {
        const masters = await getMasterDetails();
        masters.forEach((m) => {
          log({ level: "info", message: `🎨 Master: ${m.name} (${m.layoutCount} layouts)`, timestamp: Date.now() });
        });
        break;
      }
      case "listLayouts": {
        const layouts = await getAllLayouts();
        layouts.forEach((l) => {
          log({ level: "info", message: `📋 [${l.masterName}] ${l.layoutName}`, timestamp: Date.now() });
        });
        break;
      }
      case "getTheme": {
        const theme = await getThemeDetails();
        log({ level: "info", message: `🌈 Theme: ${theme.name}`, timestamp: Date.now() });
        break;
      }
      case "modifySelectedShape": {
        if (!currentSlideId) {
          log({ level: "warn", message: "No slide selected. Click a slide first.", timestamp: Date.now() });
          break;
        }
        const shapes = await getShapesOnSlide(currentSlideId);
        if (shapes.length === 0) {
          log({ level: "warn", message: "No shapes on this slide. Add a shape first (e.g. 'add rectangle shape').", timestamp: Date.now() });
          break;
        }
        // Get the last-selected or most recently added shape
        const last = shapes[shapes.length - 1];
        try {
          await setShapeFill(last.id, currentSlideId, "#9B59B6");
        } catch {
          log({ level: "warn", message: "Could not set fill (shape type may not support fill).", timestamp: Date.now() });
        }
        try {
          await setShapeGeometry(last.id, currentSlideId, { width: 200, height: 120 });
        } catch {
          log({ level: "warn", message: "Could not resize shape.", timestamp: Date.now() });
        }
        log({ level: "success", message: `Modified "${last.name || "unnamed"}" → purple fill, 200×120`, timestamp: Date.now() });
        break;
      }
      case "insertTable": {
        await addTable({
          headers: ["Name", "Score", "Grade"],
          rows: [
            ["Alice", "95", "A"],
            ["Bob", "87", "B"],
            ["Carol", "92", "A"],
          ],
        });
        log({ level: "success", message: "Inserted sample table", timestamp: Date.now() });
        break;
      }
      case "insertChart": {
        await addChart("ColumnClustered", {
          categories: ["Jan", "Feb", "Mar", "Apr", "May"],
          series: [
            { name: "Sales", values: [100, 140, 160, 190, 220] },
            { name: "Target", values: [120, 130, 150, 170, 200] },
          ],
        }, { title: "Monthly Sales Report" });
        log({ level: "success", message: "Inserted column chart", timestamp: Date.now() });
        break;
      }
      case "setSlideBackground": {
        if (currentSlideId) {
          await setSlideBackground(currentSlideId, "#1a1a2e");
          log({ level: "success", message: "Set slide background to dark", timestamp: Date.now() });
        } else {
          log({ level: "warn", message: "Select a slide first.", timestamp: Date.now() });
        }
        break;
      }
      case "autoLayout": {
        if (currentSlideId) {
          const n = await autoLayoutShapes(currentSlideId);
          log({ level: "success", message: `Auto-arranged ${n} shapes`, timestamp: Date.now() });
        } else {
          log({ level: "warn", message: "Select a slide first.", timestamp: Date.now() });
        }
        break;
      }
      case "listThemes": {
        const themes = listAvailableThemes();
        const schemes = listDesignSchemes();
        log({ level: "info", message: `🎨 Themes: ${themes.join(", ")}`, timestamp: Date.now() });
        log({ level: "info", message: `🎨 Design Schemes: ${schemes.join(", ")}`, timestamp: Date.now() });
        break;
      }
      case "applyDesign": {
        if (currentSlideId) {
          const msg = await applyDesignScheme(currentSlideId, "modern dark");
          log({ level: "success", message: msg, timestamp: Date.now() });
        } else {
          log({ level: "warn", message: "Select a slide first.", timestamp: Date.now() });
        }
        break;
      }
      default:
        log({ level: "warn", message: `Unknown action: ${action}`, timestamp: Date.now() });
    }
  } catch (err: any) {
    log({ level: "error", message: `Action failed: ${err.message || err}`, timestamp: Date.now() });
  }
}

// ── AI Badge ──────────────────────────────────────────────────────

function updateAiBadge(): void {
  if (aiEnabled) {
    aiBadge.className = "";
    aiBadge.title = "AI Mode ON — Click to disable";
    aiStatusText.className = "ai-status-on";
    aiStatusText.textContent = "AI ON";
  } else {
    aiBadge.className = "ai-off";
    aiBadge.title = "AI Mode OFF — Click to enable";
    aiStatusText.className = "ai-status-off";
    aiStatusText.textContent = "AI OFF";
  }
}

function updateApiKeyStatus(): void {
  if (hasApiKey()) {
    apiKeyStatus.className = "key-saved";
    apiKeyStatus.textContent = "✓ Saved";
  } else {
    apiKeyStatus.className = "key-missing";
    apiKeyStatus.textContent = "Not set";
  }
}

function updateHistoryCount(): void {
  const el = document.getElementById("history-count");
  if (!el) return;
  const len = getHistoryLength();
  if (len > 0) {
    el.style.display = "inline";
    el.textContent = `🧠 ${Math.floor(len / 2)} msg(s)`;
  } else {
    el.style.display = "none";
  }
}

// ── Confirm Bar ───────────────────────────────────────────────────

function showConfirmBar(count: number): void {
  const bar = document.getElementById("confirm-bar")!;
  const cnt = document.getElementById("confirm-count")!;
  bar.style.display = "flex";
  cnt.textContent = `${count} action(s) planned`;
}

function hideConfirmBar(): void {
  document.getElementById("confirm-bar")!.style.display = "none";
  pendingActions = [];
  sendBtn.disabled = false;
  sendBtn.textContent = "▶ Send";
}

async function applyPendingActions(): Promise<void> {
  if (pendingActions.length === 0) return;
  const calls = [...pendingActions]; // capture before clearing
  log({ level: "info", message: "⚡ Executing planned actions...", timestamp: Date.now() });
  hideConfirmBar();

  const results = await executePendingCalls(calls, currentSlideId);
  let ok = 0;
  for (const r of results) {
    if (r.success) { ok++; log({ level: "success", message: r.message, timestamp: Date.now() }); }
    else { log({ level: "warn", message: r.message, timestamp: Date.now() }); }
  }
  log({ level: "success", message: `✅ ${ok}/${results.length} actions applied`, timestamp: Date.now() });
}

// ── Initialize ────────────────────────────────────────────────────

async function init(): Promise<void> {
  try {
    await ensureOfficeReady();

    // Update status
    statusIndicator.textContent = "● Connected";
    statusIndicator.className = "status-connected";

    // Get current slide (fallback to first slide)
    const slide = await getSelectedSlide();
    if (slide) {
      currentSlideId = slide.id;
    } else {
      // No slide selected yet — grab the first slide as fallback
      const allSlides = await getSlides();
      if (allSlides.length > 0) {
        currentSlideId = allSlides[0].id;
        log({ level: "info", message: "Auto-selected first slide.", timestamp: Date.now() });
      }
    }

    // Get presentation info
    const info = await getPresentationInfo();
    log({ level: "info", message: `📁 ${info.title} — ${info.slideCount} slide(s)`, timestamp: Date.now() });
    log({ level: "info", message: 'Ready. Type a command or use Quick Actions below. Type "help" for commands.', timestamp: Date.now() });

    // Register event handlers — updates currentSlideId on slide change
    await registerEventHandlers(
      (entry) => log(entry),
      (newSlideId) => {
        currentSlideId = newSlideId;
      }
    );

    // ── Bind UI Events ────────────────────────────────────────────

    // API Key toggle
    apiKeyToggle.addEventListener("click", () => {
      const isVisible = apiKeyBody.style.display !== "none";
      apiKeyBody.style.display = isVisible ? "none" : "block";
    });

    // API Key save
    apiKeySave.addEventListener("click", () => {
      const key = apiKeyInput.value.trim();
      if (key) {
        setApiKey(key);
        apiKeyInput.value = "";
        apiKeyBody.style.display = "none";
        aiEnabled = true;
        updateAiBadge();
        updateApiKeyStatus();
        log({ level: "success", message: "🔑 API key saved. AI mode ON.", timestamp: Date.now() });
      }
    });

    // Restore saved key if previously set
    if (hasApiKey()) {
      aiEnabled = true;
    }
    updateAiBadge();
    updateApiKeyStatus();

    // AI badge toggle
    aiBadge.addEventListener("click", () => {
      if (!hasApiKey()) {
        log({ level: "warn", message: "Set your DeepSeek API key first (🔑 API Settings).", timestamp: Date.now() });
        apiKeyBody.style.display = "block";
        return;
      }
      aiEnabled = !aiEnabled;
      updateAiBadge();
      log({ level: "info", message: aiEnabled ? "🧠 AI mode ON — natural language commands enabled" : "📋 AI mode OFF — using regex commands", timestamp: Date.now() });
    });

    // Send button
    sendBtn.addEventListener("click", async () => {
      const cmd = commandInput.value;
      commandInput.value = "";
      await executeCommand(cmd);
      updateHistoryCount();
    });

    // Clear conversation history
    const clearHistoryBtn = document.getElementById("clear-history-btn") as HTMLButtonElement;
    const historyCount = document.getElementById("history-count") as HTMLSpanElement;
    clearHistoryBtn.addEventListener("click", () => {
      clearConversationHistory();
      updateHistoryCount();
      log({ level: "info", message: "🗑 AI conversation context cleared. Starting fresh.", timestamp: Date.now() });
    });
    updateHistoryCount();

    // Ctrl+Enter to send
    commandInput.addEventListener("keydown", async (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        const cmd = commandInput.value;
        commandInput.value = "";
        await executeCommand(cmd);
      }
    });

    // Quick action buttons
    actionButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const action = btn.dataset.action;
        if (action) handleQuickAction(action);
      });
    });

    // Quick actions toggle
    const actionsToggle = document.getElementById("actions-toggle")!;
    const actionsGrid = document.getElementById("actions-grid")!;
    actionsToggle.addEventListener("click", () => {
      const show = actionsGrid.style.display === "none";
      actionsGrid.style.display = show ? "grid" : "none";
      actionsToggle.textContent = show ? "⚡ Quick Actions ▴" : "⚡ Quick Actions ▾";
    });

    // Confirm bar buttons
    document.getElementById("confirm-apply")!.addEventListener("click", applyPendingActions);
    document.getElementById("confirm-cancel")!.addEventListener("click", () => {
      log({ level: "info", message: "❌ Cancelled.", timestamp: Date.now() });
      hideConfirmBar();
    });

  } catch (err: any) {
    statusIndicator.textContent = "● Error";
    statusIndicator.className = "status-disconnected";
    log({ level: "error", message: `Initialization failed: ${err.message || err}`, timestamp: Date.now() });
  }
}

// ── Cleanup on unload ─────────────────────────────────────────────

window.addEventListener("beforeunload", () => {
  unregisterEventHandlers();
});

// ── Start ─────────────────────────────────────────────────────────

Office.onReady(() => {
  init();
});
