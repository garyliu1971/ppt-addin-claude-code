/**
 * validateService.ts — Pre-execution validation & post-execution QA
 *
 * Step 5 & 6 of the request lifecycle:
 *   - Bounds checking (shapes within slide 960×540)
 *   - Required-args validation per tool
 *   - Overlap prediction (same-turn planned shapes)
 *   - Post-execution read-back QA
 */

export interface ValidationError {
  tool: string;
  field?: string;
  message: string;
  severity: "error" | "warn";
}

export interface ValidationResult {
  valid: boolean;
  errors: ValidationError[];
  warnings: ValidationError[];
}

// ── Slide constants ────────────────────────────────────────────────

const SLIDE_W = 960;
const SLIDE_H = 540;
const MIN_SIZE = 2;
const MAX_SIZE = 960;

const COLOR_MAP: Record<string, string> = {
  blue: "#4A90D9", red: "#E74C3C", green: "#2ECC71", yellow: "#F1C40F",
  orange: "#E67E22", purple: "#9B59B6", pink: "#E91E63", black: "#333333",
  white: "#FFFFFF", gray: "#95A5A6",
};

// ── Required fields per tool ───────────────────────────────────────

const REQUIRED_FIELDS: Record<string, string[]> = {
  add_shape: ["geometry"],
  add_image: ["url"],
  add_image_base64: ["base64"],
  add_card: ["left", "top", "heading"],
  add_text_box: ["text"],
  add_rich_text: ["paragraphs"],
  add_table: ["headers", "rows"],
  add_chart: ["chartType", "categories", "series"],
  set_slide_background: ["color"],
  apply_layout: ["layoutName"],
  delete_slide_by_index: ["index"],
  set_slide_title: ["title"],
  move_slide: ["fromIndex", "toIndex"],
  apply_theme: ["themeName"],
  apply_design_scheme: ["schemeName"],
  set_shape_format: ["shapeName"],
  set_shape_fill: ["shapeName", "color"],
  delete_shape: ["shapeName"],
  build_professional_slide: [],
  add_slide_with_title: ["title"],
};

// ── Position args across all tools ─────────────────────────────────

const POS_KEYS = ["left", "top", "width", "height"];

// ── Bounds validation ──────────────────────────────────────────────

function validateBounds(tool: string, args: Record<string, any>): ValidationError[] {
  const errs: ValidationError[] = [];
  // Only check tools that create positioned elements
  if (
    !["add_shape", "add_card", "add_text_box", "add_rich_text",
      "add_image", "add_image_base64", "add_table", "add_chart"].includes(tool)
  ) return errs;

  const left = args.left;
  const top = args.top;
  const width = args.width;
  const height = args.height;

  if (left !== undefined) {
    if (left < 0) errs.push({ tool, field: "left", message: `left=${left} is negative; clamped to 0`, severity: "warn" });
    if (left >= SLIDE_W) errs.push({ tool, field: "left", message: `left=${left} exceeds slide width ${SLIDE_W}`, severity: "error" });
  }
  if (top !== undefined) {
    if (top < 0) errs.push({ tool, field: "top", message: `top=${top} is negative; clamped to 0`, severity: "warn" });
    if (top >= SLIDE_H) errs.push({ tool, field: "top", message: `top=${top} exceeds slide height ${SLIDE_H}`, severity: "error" });
  }
  if (width !== undefined) {
    if (width < MIN_SIZE) errs.push({ tool, field: "width", message: `width=${width} too small (min ${MIN_SIZE})`, severity: "error" });
    if (width > MAX_SIZE) errs.push({ tool, field: "width", message: `width=${width} exceeds max ${MAX_SIZE}`, severity: "error" });
  }
  if (height !== undefined) {
    if (height < MIN_SIZE) errs.push({ tool, field: "height", message: `height=${height} too small (min ${MIN_SIZE})`, severity: "error" });
    if (height > MAX_SIZE) errs.push({ tool, field: "height", message: `height=${height} exceeds max ${MAX_SIZE}`, severity: "error" });
  }

  // Overflow check: left + width must fit
  if (left !== undefined && width !== undefined) {
    if (left + width > SLIDE_W) {
      errs.push({ tool, field: "width", message: `left(${left}) + width(${width}) = ${left + width} exceeds slide width ${SLIDE_W}`, severity: "error" });
    }
  }
  if (top !== undefined && height !== undefined) {
    if (top + height > SLIDE_H) {
      errs.push({ tool, field: "height", message: `top(${top}) + height(${height}) = ${top + height} exceeds slide height ${SLIDE_H}`, severity: "error" });
    }
  }

  return errs;
}

// ── Required fields validation ─────────────────────────────────────

function validateRequired(tool: string, args: Record<string, any>): ValidationError[] {
  const errs: ValidationError[] = [];
  const required = REQUIRED_FIELDS[tool];
  if (!required) return errs;

  for (const key of required) {
    if (args[key] === undefined || args[key] === null || args[key] === "") {
      errs.push({ tool, field: key, message: `"${key}" is required for ${tool}`, severity: "error" });
    }
  }
  return errs;
}

// ── Color validation ───────────────────────────────────────────────

const COLOR_KEYS = ["fillColor", "lineColor", "color", "headingColor", "subtitleColor",
  "fontColor", "textColor"];

function validateColors(tool: string, args: Record<string, any>): ValidationError[] {
  const errs: ValidationError[] = [];
  for (const key of COLOR_KEYS) {
    const val = args[key];
    if (!val || typeof val !== "string") continue;
    if (val.startsWith("#") && /^#[0-9A-Fa-f]{6}$/.test(val)) continue; // valid hex
    if (COLOR_MAP[val.toLowerCase()]) continue; // named color
    errs.push({ tool, field: key, message: `"${val}" is not a valid hex color or named color`, severity: "warn" });
  }
  return errs;
}

// ── Overlap prediction (same-batch planned shapes) ─────────────────

interface PlannedShape {
  tool: string;
  left?: number; top?: number; width?: number; height?: number;
}

function predictOverlaps(shapes: PlannedShape[]): ValidationError[] {
  const errs: ValidationError[] = [];
  const positioned = shapes.filter(
    s => s.left !== undefined && s.top !== undefined &&
         s.width !== undefined && s.height !== undefined
  );

  for (let i = 0; i < positioned.length; i++) {
    for (let j = i + 1; j < positioned.length; j++) {
      const a = positioned[i], b = positioned[j];
      const aR = a.left! + a.width!;
      const aB = a.top! + a.height!;
      const bR = b.left! + b.width!;
      const bB = b.top! + b.height!;

      // Check if they overlap
      if (!(a.left! < bR && aR > b.left! && a.top! < bB && aB > b.top!)) continue;

      // Calculate overlap area
      const ox = Math.max(0, Math.min(aR, bR) - Math.max(a.left!, b.left!));
      const oy = Math.max(0, Math.min(aB, bB) - Math.max(a.top!, b.top!));
      const overlapArea = ox * oy;
      const minArea = Math.min(a.width! * a.height!, b.width! * b.height!);
      const overlapPct = minArea > 0 ? overlapArea / minArea : 0;

      const msg = `Shape "${a.tool}" overlaps ${(overlapPct * 100).toFixed(0)}% with "${b.tool}" at (${a.left},${a.top})↔(${b.left},${b.top})`;

      // Severe overlap (>30% of smaller shape) → error
      if (overlapPct > 0.3) {
        errs.push({ tool: a.tool, message: msg, severity: "error" });
      } else {
        errs.push({ tool: a.tool, message: msg, severity: "warn" });
      }
    }
  }
  return errs;
}

// ── Main validation entry ──────────────────────────────────────────

/**
 * Validate a single tool call BEFORE execution.
 * Returns structured errors/warnings.
 *
 * @param tool - tool function name
 * @param args - parsed arguments
 * @param plannedShapes - all shapes planned in this batch (for overlap check)
 */
export function validateToolCall(
  tool: string,
  args: Record<string, any>,
  plannedShapes?: PlannedShape[]
): ValidationResult {
  const errors: ValidationError[] = [];
  const warnings: ValidationError[] = [];

  // 1) Required fields
  for (const e of validateRequired(tool, args)) {
    if (e.severity === "error") errors.push(e);
    else warnings.push(e);
  }

  // 2) Bounds
  for (const e of validateBounds(tool, args)) {
    if (e.severity === "error") errors.push(e);
    else warnings.push(e);
  }

  // 3) Colors
  for (const e of validateColors(tool, args)) {
    warnings.push(e);
  }

  // 4) Overlap (only for non-empty batch)
  if (plannedShapes && plannedShapes.length > 1) {
    // Add current tool as a planned shape
    const currentShape: PlannedShape = {
      tool,
      left: args.left, top: args.top, width: args.width, height: args.height,
    };
    const allShapes = [...plannedShapes, currentShape];
    for (const e of predictOverlaps(allShapes)) {
      warnings.push(e);
    }
  }

  return { valid: errors.length === 0, errors, warnings };
}

/**
 * Validate a batch of tool calls. Collects all planned shapes
 * and runs overlap detection across the batch.
 */
export function validateBatch(
  calls: { name: string; args: Record<string, any> }[]
): ValidationResult[] {
  const plannedShapes: PlannedShape[] = [];
  const results: ValidationResult[] = [];

  for (const call of calls) {
    const result = validateToolCall(call.name, call.args, [...plannedShapes]);
    results.push(result);

    // Track this call as a planned shape for subsequent overlap checks
    if (call.args.left !== undefined && call.args.top !== undefined) {
      plannedShapes.push({
        tool: call.name,
        left: call.args.left,
        top: call.args.top,
        width: call.args.width,
        height: call.args.height,
      });
    }
  }

  return results;
}

// ── Post-execution QA ──────────────────────────────────────────────

import { getShapesOnSlide } from "./pptApi";

/**
 * Read-back QA: after executing tools against a slide,
 * verify that the expected number of shapes are present.
 *
 * @param slideId - target slide ID
 * @param expectedMin - minimum expected shape count
 * @param toolNames - names of tools executed (for context)
 */
export async function readBackQA(
  slideId: string,
  expectedMin: number,
  toolNames: string[]
): Promise<{ ok: boolean; message: string }> {
  try {
    const shapes = await getShapesOnSlide(slideId);
    const count = shapes.length;
    if (count < expectedMin) {
      return {
        ok: false,
        message: `⚠️ QA: Expected ≥${expectedMin} shapes on slide, found ${count}. Tools: ${toolNames.join(", ")}`,
      };
    }
    return {
      ok: true,
      message: `✅ QA: ${count} shapes on slide after ${toolNames.join(", ")}`,
    };
  } catch {
    return { ok: false, message: "⚠️ QA: Could not read back shapes from slide" };
  }
}
