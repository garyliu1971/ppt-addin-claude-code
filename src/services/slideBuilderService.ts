/**
 * slideBuilderService.ts — Professional slide builder (data-driven)
 *
 * Provides a high-level `buildSlide(schema)` API that creates polished,
 * magazine-style slides with background, eyebrow, title, description,
 * multi-column cards, insight bars, and footer — all from a JSON schema.
 *
 * Slide dimensions: 960 x 540 pt (widescreen 16:9).
 *
 * Design tokens are customisable per schema or fall back to defaults.
 */

import { runPPT, getSelectedSlide } from "./pptApi";

// ── Types ──────────────────────────────────────────────────────────

export interface DesignTokens {
  bg: string;           // slide background
  accent: string;       // highlight bars, eyebrow marker
  cardBg: string;       // card background
  cardBorder: string;   // card border color
  white: string;        // primary text
  muted: string;        // secondary text
  headingFont: string;  // e.g. "Arial Black"
  bodyFont: string;     // e.g. "Arial"
}

export const DEFAULT_DARK_TOKENS: DesignTokens = {
  bg: "0A1A2F", accent: "27E07C", cardBg: "0F2740", cardBorder: "12314F",
  white: "FFFFFF", muted: "9FB8CE",
  headingFont: "Arial Black", bodyFont: "Arial",
};

/** A single player / item inside a card */
export interface CardItem {
  name: string;
  description: string;
}

/** A column card (e.g. team card) */
export interface CardColumn {
  title: string;
  subtitle?: string;
  items: CardItem[];
}

export interface ProfessionalSlideSchema {
  /** Design color tokens (defaults to dark theme) */
  tokens?: Partial<DesignTokens>;
  /** Small label above title, with accent marker bar */
  eyebrow?: string;
  /** Main slide title */
  title: string;
  /** Subtitle / description paragraph */
  description?: string;
  /** Data-driven column cards (3 max recommended) */
  columns?: CardColumn[];
  /** Bottom insight / outlook section */
  insight?: {
    label: string;
    body: string;
  };
  /** Footer left & right text */
  footer?: { left: string; right: string };
}

// ── Internal constants ───────────────────────────────────────────────

/** Slide width/height in pt */
const SW = 960, SH = 540;

// ── Inline helpers (no runPPT — use within a single batch) ─────────

function addRect(
  shapes: PowerPoint.ShapeCollection,
  name: string, x: number, y: number, w: number, h: number,
  fillColor: string | null, t: DesignTokens,
  opts?: { round?: boolean; border?: boolean; accentBar?: boolean }
) {
  const r = shapes.addGeometricShape(
    (opts?.round ? "RoundRectangle" : "Rectangle") as any
  );
  r.left = x; r.top = y; r.width = w; r.height = h;
  (r as any).name = name;
  if (opts?.accentBar) {
    r.fill.setSolidColor(t.accent);
  } else if (fillColor) {
    r.fill.setSolidColor(fillColor);
  } else {
    r.fill.clear();
  }
  if (opts?.border) {
    r.lineFormat.visible = true;
    (r.lineFormat as any).color = t.cardBorder;
    r.lineFormat.weight = 1;
  } else {
    r.lineFormat.visible = false;
  }
  return r;
}

function addTextBox(
  shapes: PowerPoint.ShapeCollection,
  name: string, x: number, y: number, w: number, h: number,
  text: string
) {
  const tb = shapes.addTextBox(text);
  tb.left = x; tb.top = y; tb.width = w; tb.height = h;
  (tb as any).name = name;
  tb.fill.clear();
  tb.lineFormat.visible = false;
  return tb;
}

// ── Text style to apply after sync ──────────────────────────────────

interface TextStyle {
  name: string;
  fontSize: number;
  fontColor: string;
  fontName: string;
  bold?: boolean;
  align?: string;
  vAlign?: string;
}

// ── Main builder (two-phase: create → sync → style → sync) ──────────

export async function buildProfessionalSlide(
  schema: ProfessionalSlideSchema,
  slideId?: string,
): Promise<{ slideId: string; shapeCount: number }> {
  let targetId = slideId;
  if (!targetId) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected and no slideId provided.");
    targetId = sel.id;
  }

  const t: DesignTokens = { ...DEFAULT_DARK_TOKENS, ...(schema.tokens || {}) };
  const textStyles: TextStyle[] = [];

  await runPPT(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(targetId).shapes;

    // ── Phase 1: Create all shapes ──

    addRect(shapes, "BG", 0, 0, SW, SH, t.bg, t);

    if (schema.eyebrow) {
      addRect(shapes, "EyebrowMarker", 50, 48, 25, 7, null, t, { accentBar: true });
      addTextBox(shapes, "EyebrowText", 84, 37, 648, 29, schema.eyebrow);
      textStyles.push({ name: "EyebrowText", fontSize: 12, fontColor: t.accent, fontName: t.headingFont, bold: true });
    }

    const titleY = schema.eyebrow ? 66 : 48;
    addTextBox(shapes, "Title", 50, titleY, 857, 55, schema.title);
    textStyles.push({ name: "Title", fontSize: 28, fontColor: t.white, fontName: t.headingFont, bold: true });

    let nextY = titleY + 58;
    if (schema.description) {
      addTextBox(shapes, "Description", 50, nextY, 857, 44, schema.description);
      textStyles.push({ name: "Description", fontSize: 14, fontColor: t.muted, fontName: t.bodyFont, vAlign: "top" });
      nextY += 44;
    }

    const cols = schema.columns || [];
    if (cols.length > 0) {
      const cardY = nextY + 14, gap = 22, margin = 50;
      const cardW = (SW - margin * 2 - gap * (cols.length - 1)) / cols.length;
      const cardH = 198;

      for (let i = 0; i < cols.length; i++) {
        const col = cols[i], cx = margin + i * (cardW + gap), pad = 20;

        addRect(shapes, `Card${i + 1}_BG`, cx, cardY, cardW, cardH, t.cardBg, t, { round: true, border: true });
        addRect(shapes, `Card${i + 1}_Bar`, cx + pad, cardY + 18, 17, 5, null, t, { accentBar: true });
        addTextBox(shapes, `Card${i + 1}_Title`, cx + pad, cardY + 26, cardW - pad * 2, 26, col.title);
        textStyles.push({ name: `Card${i + 1}_Title`, fontSize: 16, fontColor: t.white, fontName: t.headingFont, bold: true });

        if (col.subtitle) {
          addTextBox(shapes, `Card${i + 1}_Sub`, cx + pad, cardY + 52, cardW - pad * 2, 20, col.subtitle);
          textStyles.push({ name: `Card${i + 1}_Sub`, fontSize: 9.5, fontColor: t.accent, fontName: t.headingFont, bold: true });
        }

        const itemStartY = cardY + (col.subtitle ? 80 : 62);
        for (let j = 0; j < col.items.length; j++) {
          const iy = itemStartY + j * 39;
          addTextBox(shapes, `C${i + 1}_I${j + 1}_Name`, cx + pad, iy, cardW - pad * 2, 19, col.items[j].name);
          textStyles.push({ name: `C${i + 1}_I${j + 1}_Name`, fontSize: 12, fontColor: t.white, fontName: t.headingFont, bold: true });
          addTextBox(shapes, `C${i + 1}_I${j + 1}_Desc`, cx + pad, iy + 19, cardW - pad * 2, 18, col.items[j].description);
          textStyles.push({ name: `C${i + 1}_I${j + 1}_Desc`, fontSize: 9, fontColor: t.muted, fontName: t.bodyFont });
        }
      }
      nextY = cardY + cardH + 14;
    }

    if (schema.insight) {
      const barY = Math.min(nextY + 6, SH - 146);
      addRect(shapes, "Insight_BG", 50, barY, 857, 97, t.cardBg, t, { round: true, border: true });
      addRect(shapes, "Insight_Bar", 50, barY, 6.5, 97, null, t, { accentBar: true });
      addTextBox(shapes, "Insight_Label", 71, barY + 12, 394, 24, schema.insight.label);
      textStyles.push({ name: "Insight_Label", fontSize: 12.5, fontColor: t.accent, fontName: t.headingFont, bold: true });
      addTextBox(shapes, "Insight_Body", 71, barY + 37, 803, 54, schema.insight.body);
      textStyles.push({ name: "Insight_Body", fontSize: 11, fontColor: t.white, fontName: t.bodyFont, vAlign: "top" });
    }

    if (schema.footer) {
      addTextBox(shapes, "Footer_L", 50, SH - 33, 432, 22, schema.footer.left);
      textStyles.push({ name: "Footer_L", fontSize: 9, fontColor: t.muted, fontName: t.bodyFont });
      addTextBox(shapes, "Footer_R", 504, SH - 33, 406, 22, schema.footer.right);
      textStyles.push({ name: "Footer_R", fontSize: 9, fontColor: t.muted, fontName: t.bodyFont, align: "right" });
    }

    // ── CRITICAL: sync after creating all shapes ──
    await ctx.sync();

    // ── Phase 2: Load all shapes and apply text styles by name ──
    const slideShapes = ctx.presentation.slides.getItem(targetId).shapes;
    slideShapes.load("items/name, items/textFrame/textRange/font, items/textFrame/textRange/paragraphFormat, items/textFrame/verticalAlignment");
    await ctx.sync();

    const shapeMap = new Map<string, PowerPoint.Shape>();
    for (const s of slideShapes.items) {
      const sName = (s as any).name || "";
      if (sName) shapeMap.set(sName, s);
    }

    for (const style of textStyles) {
      try {
        const s = shapeMap.get(style.name);
        if (!s) continue;

        const tr = s.textFrame.textRange;
        tr.font.name = style.fontName;
        tr.font.size = style.fontSize;
        tr.font.color = style.fontColor;
        if (style.bold !== undefined) tr.font.bold = style.bold;
        if (style.align) {
          const aMap: Record<string, string> = { left: "Left", center: "Center", right: "Right" };
          (tr.paragraphFormat as any).horizontalAlignment = aMap[style.align] || "Left";
        }
        if (style.vAlign) {
          const vMap: Record<string, string> = { top: "Top", middle: "Middle", bottom: "Bottom" };
          (s.textFrame as any).verticalAlignment = vMap[style.vAlign] || "Middle";
        }
        await ctx.sync();
      } catch {
        // best-effort: skip shapes that don't support text
      }
    }
  });

  return { slideId: targetId, shapeCount: 1 + textStyles.length };
}

function findShapeByName(shapes: PowerPoint.ShapeCollection, name: string): PowerPoint.Shape | null {
  try { return (shapes as any).getItem(name); } catch { return null; }
}

// ── Pre-built demo: NBA Focus Teams ────────────────────────────────

export function getNBADemoData(): ProfessionalSlideSchema {
  return {
    eyebrow: "NBA 焦点球队 · STAR POWER",
    title: "Warriors · Lakers · Cavaliers 的当家球星",
    description:
      "2025-26 赛季落幕,三支传统豪门在新老交替中重新洗牌——这是他们的核心球星,以及通往 2026-27 的展望。",
    columns: [
      {
        title: "Warriors 勇士",
        subtitle: "2025-26 · 37–45 · 未进季后赛",
        items: [
          { name: "Stephen Curry 库里", description: "四届总冠军、球队基石,38 岁仍是进攻核心" },
          { name: "Jimmy Butler 巴特勒", description: "2025 年五队交易加盟,季后赛型攻防领袖" },
          { name: "Draymond Green 格林", description: "防守中枢与组织核心,王朝拼图" },
        ],
      },
      {
        title: "Lakers 湖人",
        subtitle: "2025-26 · 53–29 · 太平洋分区冠军",
        items: [
          { name: "Luka Dončić 东契奇", description: "2025 年重磅交易加盟,新一代当家核心" },
          { name: "LeBron James 詹姆斯", description: "第 24 季老将,已告知 2026-27 将转投他队" },
          { name: "Austin Reaves 里夫斯", description: "稳定的后场得分与串联点" },
        ],
      },
      {
        title: "Cavaliers 骑士",
        subtitle: "2025-26 · 52–30 · 东部劲旅",
        items: [
          { name: "Donovan Mitchell 米切尔", description: "球队箭头,进攻发起点" },
          { name: "Evan Mobley 莫布利", description: "内线防守屏障,攻防一体潜力核心" },
          { name: "James Harden 哈登", description: "2026 年截止日加盟,即战力控卫" },
        ],
      },
    ],
    insight: {
      label: "展望 OUTLOOK · 2026-27",
      body: "湖人以 27 岁的东契奇为核心重建、送别詹姆斯时代;勇士围绕 38 岁的库里追逐最后的冠军窗口,并传出追求安东尼·戴维斯与勒布朗的运作;骑士押注哈登即战力,力争在东部更进一步。东契奇、莫布利等新生代与库里、哈登等老将的碰撞,将定义下一阶段的联盟格局。",
    },
    footer: { left: "NBA · 2025-26 SEASON", right: "焦点球队 · Star Power" },
  };
}

/**
 * Shortcut: build the NBA demo slide on the currently selected slide.
 */
export async function buildNBADemoSlide(slideId?: string) {
  return buildProfessionalSlide(getNBADemoData(), slideId);
}

/**
 * Build a custom professional slide from user-provided data.
 * Accepts the full schema plus an optional target slideId.
 */
export async function buildCustomSlide(
  data: ProfessionalSlideSchema,
  slideId?: string,
) {
  return buildProfessionalSlide(data, slideId);
}
