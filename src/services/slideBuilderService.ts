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

// ── Internal helpers ───────────────────────────────────────────────

interface RectSpec {
  x: number; y: number; w: number; h: number;
  fill?: string; round?: boolean; border?: boolean;
  accentBar?: boolean;
}

interface TextSpec {
  x: number; y: number; w: number; h: number;
  text: string;
  size: number; color: string; font: string;
  bold?: boolean; align?: string; vAlign?: string;
}

/** Slide width/height in pt */
const SW = 960, SH = 540;

function rect(slideId: string, name: string, s: RectSpec, t: DesignTokens) {
  return runPPT(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(slideId).shapes;
    const r = shapes.addGeometricShape(
      (s.round ? "RoundRectangle" : "Rectangle") as any
    );
    r.left = s.x; r.top = s.y; r.width = s.w; r.height = s.h;
    (r as any).name = name;
    if (s.accentBar) {
      r.fill.setSolidColor(t.accent);
    } else if (s.fill) {
      r.fill.setSolidColor(s.fill);
    } else {
      r.fill.clear();
    }
    if (s.border) {
      r.lineFormat.visible = true;
      (r.lineFormat as any).color = t.cardBorder;
      r.lineFormat.weight = 1;
    } else {
      r.lineFormat.visible = false;
    }
    await ctx.sync();
    return r;
  });
}

function textBox(slideId: string, name: string, s: TextSpec) {
  return runPPT(async (ctx) => {
    const shapes = ctx.presentation.slides.getItem(slideId).shapes;
    const tb = shapes.addTextBox(s.text);
    tb.left = s.x; tb.top = s.y; tb.width = s.w; tb.height = s.h;
    (tb as any).name = name;
    tb.fill.clear();
    tb.lineFormat.visible = false;
    await ctx.sync();

    const tr = tb.textFrame.textRange;
    tr.font.name = s.font;
    tr.font.size = s.size;
    tr.font.color = s.color;
    if (s.bold !== undefined) tr.font.bold = s.bold;
    if (s.align) {
      const map: Record<string, string> = {
        left: "Left", center: "Center", right: "Right", justify: "Justify",
      };
      (tr.paragraphFormat as any).horizontalAlignment = map[s.align] || "Left";
    }
    if (s.vAlign) {
      const map: Record<string, string> = {
        top: "Top", middle: "Middle", bottom: "Bottom",
      };
      (tb.textFrame as any).verticalAlignment = map[s.vAlign] || "Middle";
    }
    await ctx.sync();
    return tb;
  });
}

// ── Main builder ───────────────────────────────────────────────────

/**
 * Build a professional slide from a JSON schema.
 *
 * @param schema - The slide content & styling
 * @param slideId - Optional: target a specific slide; defaults to selected slide
 * @returns count of shapes created
 */
export async function buildProfessionalSlide(
  schema: ProfessionalSlideSchema,
  slideId?: string,
): Promise<{ slideId: string; shapeCount: number }> {
  // Resolve target slide
  let targetId = slideId;
  if (!targetId) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected and no slideId provided.");
    targetId = sel.id;
  }

  const t: DesignTokens = { ...DEFAULT_DARK_TOKENS, ...(schema.tokens || {}) };

  // ── 1) Background ──
  await rect(targetId, "BG", { x: 0, y: 0, w: SW, h: SH, fill: t.bg }, t);

  // ── 2) Eyebrow ──
  if (schema.eyebrow) {
    await rect(targetId, "EyebrowMarker", { x: 50, y: 48, w: 25, h: 7, accentBar: true }, t);
    await textBox(targetId, "EyebrowText", {
      x: 84, y: 37, w: 648, h: 29,
      text: schema.eyebrow,
      size: 12, color: t.accent, font: t.headingFont, bold: true,
    });
  }

  // ── 3) Title ──
  const titleY = schema.eyebrow ? 66 : 48;
  await textBox(targetId, "Title", {
    x: 50, y: titleY, w: 857, h: 55,
    text: schema.title,
    size: 28, color: t.white, font: t.headingFont, bold: true,
  });

  // ── 4) Description ──
  let nextY = titleY + 58;
  if (schema.description) {
    await textBox(targetId, "Description", {
      x: 50, y: nextY, w: 857, h: 44,
      text: schema.description,
      size: 14, color: t.muted, font: t.bodyFont, vAlign: "top",
    });
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

      // Card background
      await rect(targetId, `Card${i + 1}_BG`, {
        x: cx, y: cardY, w: cardW, h: cardH,
        fill: t.cardBg, round: true, border: true,
      }, t);

      // Accent bar top
      await rect(targetId, `Card${i + 1}_Bar`, {
        x: cx + pad, y: cardY + 18, w: 17, h: 5, accentBar: true,
      }, t);

      // Team/column title
      await textBox(targetId, `Card${i + 1}_Title`, {
        x: cx + pad, y: cardY + 26, w: cardW - pad * 2, h: 26,
        text: col.title,
        size: 16, color: t.white, font: t.headingFont, bold: true,
      });

      // Subtitle
      if (col.subtitle) {
        await textBox(targetId, `Card${i + 1}_Sub`, {
          x: cx + pad, y: cardY + 52, w: cardW - pad * 2, h: 20,
          text: col.subtitle,
          size: 9.5, color: t.accent, font: t.headingFont, bold: true,
        });
      }

      // Items (name + description pairs)
      const itemStartY = cardY + (col.subtitle ? 80 : 62);
      const itemGap = 39;
      for (let j = 0; j < col.items.length; j++) {
        const iy = itemStartY + j * itemGap;
        await textBox(targetId, `C${i + 1}_I${j + 1}_Name`, {
          x: cx + pad, y: iy, w: cardW - pad * 2, h: 19,
          text: col.items[j].name,
          size: 12, color: t.white, font: t.headingFont, bold: true,
        });
        await textBox(targetId, `C${i + 1}_I${j + 1}_Desc`, {
          x: cx + pad, y: iy + 19, w: cardW - pad * 2, h: 18,
          text: col.items[j].description,
          size: 9, color: t.muted, font: t.bodyFont,
        });
      }
    }

    nextY = cardY + cardH + 14;
  }

  // ── 6) Insight / outlook bar ──
  if (schema.insight) {
    const barY = Math.min(nextY + 6, SH - 146); // prevent overflow
    const barH = 97;
    await rect(targetId, "Insight_BG", {
      x: 50, y: barY, w: 857, h: barH,
      fill: t.cardBg, round: true, border: true,
    }, t);

    // Left accent stripe
    await rect(targetId, "Insight_Bar", {
      x: 50, y: barY, w: 6.5, h: barH, accentBar: true,
    }, t);

    await textBox(targetId, "Insight_Label", {
      x: 71, y: barY + 12, w: 394, h: 24,
      text: schema.insight.label,
      size: 12.5, color: t.accent, font: t.headingFont, bold: true,
    });

    await textBox(targetId, "Insight_Body", {
      x: 71, y: barY + 37, w: 803, h: 54,
      text: schema.insight.body,
      size: 11, color: t.white, font: t.bodyFont, vAlign: "top",
    });
  }

  // ── 7) Footer ──
  if (schema.footer) {
    await textBox(targetId, "Footer_L", {
      x: 50, y: SH - 33, w: 432, h: 22,
      text: schema.footer.left,
      size: 9, color: t.muted, font: t.bodyFont,
    });
    await textBox(targetId, "Footer_R", {
      x: 504, y: SH - 33, w: 406, h: 22,
      text: schema.footer.right,
      size: 9, color: t.muted, font: t.bodyFont, align: "right",
    });
  }

  // Count shapes created (approximate)
  let count = 1; // BG
  if (schema.eyebrow) count += 2;
  count += 1; // title
  if (schema.description) count += 1;
  if (cols.length) count += cols.length * (6 + cols.reduce((s, c) => s + c.items.length * 2, 0)); // cards
  if (schema.insight) count += 4;
  if (schema.footer) count += 2;

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
