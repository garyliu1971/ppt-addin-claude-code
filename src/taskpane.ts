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
import { LogEntry, ensureOfficeReady, getSlides, getSelectedSlide, getShapesOnSlide, getPresentationInfo } from "./services/pptApi";
import { addShape, addTextBox, setShapeFill, setShapeText, deleteShape, applyStyleToAllShapes, setShapeGeometry } from "./services/shapeService";
import { addTable, addChart, ChartType, ChartData } from "./services/chartTableService";
import { getMasterDetails, applyLayoutToSlide, findLayoutByName, setSlideBackground, getThemeDetails, addSlide, deleteSlide, getAllLayouts } from "./services/masterLayoutThemeService";
import { registerEventHandlers, unregisterEventHandlers } from "./services/eventService";

// ── State ─────────────────────────────────────────────────────────

let currentSlideId: string | null = null;

// ── DOM Elements ──────────────────────────────────────────────────

const commandInput = document.getElementById("command-input") as HTMLTextAreaElement;
const sendBtn = document.getElementById("send-btn") as HTMLButtonElement;
const commandHistory = document.getElementById("command-history") as HTMLDivElement;
const outputContent = document.getElementById("output-content") as HTMLPreElement;
const statusIndicator = document.getElementById("status-indicator") as HTMLSpanElement;
const actionButtons = document.querySelectorAll<HTMLButtonElement>(".action-btn");

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

// ── Command Parser ────────────────────────────────────────────────

/**
 * Simple natural-language command parser.
 * Parses user input and delegates to the appropriate service.
 */
async function executeCommand(input: string): Promise<void> {
  const cmd = input.trim();
  if (!cmd) return;

  // Add user command to history
  const userDiv = document.createElement("div");
  userDiv.className = "command-entry user";
  userDiv.textContent = "> " + cmd;
  commandHistory.appendChild(userDiv);
  commandHistory.scrollTop = commandHistory.scrollHeight;

  log({ level: "info", message: `Executing: "${cmd}"`, timestamp: Date.now() });

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
    else if (/add (a |an )?(\w+) shape/i.test(cmd)) {
      const shapeMatch = cmd.match(/add (a |an )?(\w+) shape/i);
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
      const sizeMatch = cmd.match(/(\d+)\s*[x×]\s*(\d+)/);
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
      const newSlide = await addSlide();
      log({ level: "success", message: `Added new slide (id: ${newSlide.id})`, timestamp: Date.now() });
    }
    else if (/delete (this )?slide/i.test(cmd)) {
      if (currentSlideId) {
        await deleteSlide(currentSlideId);
        currentSlideId = null;
        log({ level: "success", message: "Deleted current slide", timestamp: Date.now() });
      }
    }
    else if (/show theme/i.test(cmd) || /get theme/i.test(cmd)) {
      const theme = await getThemeDetails();
      log({ level: "info", message: `Current theme: ${theme.name}`, timestamp: Date.now() });
    }

    // ── Info / Help ───────────────────────────────────────────────

    else if (/help/i.test(cmd)) {
      log({ level: "info", message: `Available commands:
• add [rectangle|oval|triangle|diamond|arrow|heart|star5] shape
• add text box "your text"
• make all shapes [blue|red|green|yellow|orange|purple]
• fill shape "name" with [color]
• add table
• add [bar|column|pie|line|area] chart
• apply layout "name"
• set background [color]
• add slide / delete slide
• resize shape W×H
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
      default:
        log({ level: "warn", message: `Unknown action: ${action}`, timestamp: Date.now() });
    }
  } catch (err: any) {
    log({ level: "error", message: `Action failed: ${err.message || err}`, timestamp: Date.now() });
  }
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

    // Register event handlers
    await registerEventHandlers((entry) => log(entry));

    // ── Bind UI Events ────────────────────────────────────────────

    // Send button
    sendBtn.addEventListener("click", async () => {
      const cmd = commandInput.value;
      commandInput.value = "";
      await executeCommand(cmd);
    });

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
