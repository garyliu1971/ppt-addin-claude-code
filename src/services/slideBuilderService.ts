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
  text: string,
  opts: { size: number; color: string; font: string; bold?: boolean; align?: string; vAlign?: string }
) {
  const tb = shapes.addTextBox(text);
  tb.left = x; tb.top = y; tb.width = w; tb.height = h;
  (tb as any).name = name;
  tb.fill.clear();
  tb.lineFormat.visible = false;

  const tr = tb.textFrame.textRange;
  tr.font.name = opts.font;
  tr.font.size = opts.size;
  tr.font.color = opts.color;
  if (opts.bold !== undefined) tr.font.bold = opts.bold;
  if (opts.align) {
    const map: Record<string, string> = {
      left: "Left", center: "Center", right: "Right", justify: "Justify",
    };
    (tr.paragraphFormat as any).horizontalAlignment = map[opts.align] || "Left";
  }
  if (opts.vAlign) {
    const map: Record<string, string> = {
      top: "Top", middle: "Middle", bottom: "Bottom",
    };
    (tb.textFrame as any).verticalAlignment = map[opts.vAlign] || "Middle";
  }
  return tb;
}

// ── Main builder (single PowerPoint.run batch) ─────────────────────

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

  let count = 0;

  await runPPT(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(targetId).shapes;

    // ── 1) Background ──
    addRect(shapes, "BG", 0, 0, SW, SH, t.bg, t);
    count++;

    // ── 2) Eyebrow ──
    if (schema.eyebrow) {
      addRect(shapes, "EyebrowMarker", 50, 48, 25, 7, null, t, { accentBar: true });
      addTextBox(shapes, "EyebrowText", 84, 37, 648, 29, schema.eyebrow,
        { size: 12, color: t.accent, font: t.headingFont, bold: true });
      count += 2;
    }

    // ── 3) Title ──
    const titleY = schema.eyebrow ? 66 : 48;
    addTextBox(shapes, "Title", 50, titleY, 857, 55, schema.title,
      { size: 28, color: t.white, font: t.headingFont, bold: true });
    count++;

    // ── 4) Description ──
    let nextY = titleY + 58;
    if (schema.description) {
      addTextBox(shapes, "Description", 50, nextY, 857, 44, schema.description,
        { size: 14, color: t.muted, font: t.bodyFont, vAlign: "top" });
      count++;
      nextY += 44;
    }

    // ── 5) Column cards ──
    const cols = schema.columns || [];
    if (cols.length > 0) {
      const cardY = nextY + 14;
      const gap = 22;
      const margin = 50;
      const totalGap = gap * (cols.length - 1);
      const cardW = (SW - margin * 2 - totalGap) / cols.length;
      const cardH = 198;

      for (let i = 0; i < cols.length; i++) {
        const col = cols[i];
        const cx = margin + i * (cardW + gap);
        const pad = 20;

        addRect(shapes, `Card${i + 1}_BG`, cx, cardY, cardW, cardH, t.cardBg, t,
          { round: true, border: true });
        count++;

        addRect(shapes, `Card${i + 1}_Bar`, cx + pad, cardY + 18, 17, 5, null, t,
          { accentBar: true });
        count++;

        addTextBox(shapes, `Card${i + 1}_Title`, cx + pad, cardY + 26, cardW - pad * 2, 26, col.title,
          { size: 16, color: t.white, font: t.headingFont, bold: true });
        count++;

        if (col.subtitle) {
          addTextBox(shapes, `Card${i + 1}_Sub`, cx + pad, cardY + 52, cardW - pad * 2, 20, col.subtitle,
            { size: 9.5, color: t.accent, font: t.headingFont, bold: true });
          count++;
        }

        const itemStartY = cardY + (col.subtitle ? 80 : 62);
        const itemGap = 39;
        for (let j = 0; j < col.items.length; j++) {
          const iy = itemStartY + j * itemGap;
          addTextBox(shapes, `C${i + 1}_I${j + 1}_Name`, cx + pad, iy, cardW - pad * 2, 19, col.items[j].name,
            { size: 12, color: t.white, font: t.headingFont, bold: true });
          addTextBox(shapes, `C${i + 1}_I${j + 1}_Desc`, cx + pad, iy + 19, cardW - pad * 2, 18, col.items[j].description,
            { size: 9, color: t.muted, font: t.bodyFont });
          count += 2;
        }
      }

      nextY = cardY + cardH + 14;
    }

    // ── 6) Insight / outlook bar ──
    if (schema.insight) {
      const barY = Math.min(nextY + 6, SH - 146);
      const barH = 97;
      addRect(shapes, "Insight_BG", 50, barY, 857, barH, t.cardBg, t,
        { round: true, border: true });
      addRect(shapes, "Insight_Bar", 50, barY, 6.5, barH, null, t,
        { accentBar: true });
      addTextBox(shapes, "Insight_Label", 71, barY + 12, 394, 24, schema.insight.label,
        { size: 12.5, color: t.accent, font: t.headingFont, bold: true });
      addTextBox(shapes, "Insight_Body", 71, barY + 37, 803, 54, schema.insight.body,
        { size: 11, color: t.white, font: t.bodyFont, vAlign: "top" });
      count += 4;
    }

    // ── 7) Footer ──
    if (schema.footer) {
      addTextBox(shapes, "Footer_L", 50, SH - 33, 432, 22, schema.footer.left,
        { size: 9, color: t.muted, font: t.bodyFont });
      addTextBox(shapes, "Footer_R", 504, SH - 33, 406, 22, schema.footer.right,
        { size: 9, color: t.muted, font: t.bodyFont, align: "right" });
      count += 2;
    }

    await ctx.sync();
  });

  return { slideId: targetId, shapeCount: count };
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
