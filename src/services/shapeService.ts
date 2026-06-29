/**
 * shapeService.ts — Shape creation and modification
 */

import { runPPT, getSelectedSlide } from "./pptApi";

export interface ShapeOptions {
  left?: number; top?: number; width?: number; height?: number;
  fillColor?: string; lineColor?: string; lineWeight?: number;
  fontSize?: number;
}

export async function addImage(imageUrl: string, opts: ShapeOptions = {}): Promise<any> {
  const slide = await getSelectedSlide();
  if (!slide) throw new Error("No slide selected.");

  // Fetch image and convert to base64 (try direct, then CORS proxy)
  let base64: string;
  try {
    // Try direct fetch first
    let resp = await fetch(imageUrl, { signal: AbortSignal.timeout(8000) });
    // Fallback: CORS proxy
    if (!resp.ok) {
      const proxyUrl = `https://corsproxy.io/?url=${encodeURIComponent(imageUrl)}`;
      resp = await fetch(proxyUrl, { signal: AbortSignal.timeout(8000) });
    }
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const blob = await resp.blob();
    base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch {
    throw new Error(`Failed to load image from: ${imageUrl}`);
  }

  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;
    // No native addPicture in PowerPoint JS — use rectangle with picture fill
    const rect = shapes.addGeometricShape("Rectangle" as any);
    rect.left = opts.left ?? 100; rect.top = opts.top ?? 100;
    rect.width = opts.width ?? 300; rect.height = opts.height ?? 200;
    (rect as any).name = `Image_${Date.now()}`;

    // Try picture fill via setImage (runtime API)
    try {
      await context.sync();
      const fill: any = rect.fill;
      fill.setImage(base64);
      await context.sync();
    } catch {
      // Picture fill failed — add a text label instead
      const label = shapes.addTextBox(`[Image]\n${imageUrl.slice(0, 50)}...`);
      label.left = rect.left + 10;
      label.top = rect.top + rect.height / 2 - 20;
      label.width = rect.width - 20;
      label.height = 40;
      label.textFrame.load("textRange/font"); await context.sync();
      label.textFrame.textRange.font.size = 9;
      label.textFrame.textRange.font.color = "#888888";
    }

    await context.sync();
    rect.load("id, name, type"); await context.sync();
    return rect;
  });
}

export async function addShape(geometry: string, opts: ShapeOptions = {}): Promise<any> {
  const slide = await getSelectedSlide();
  if (!slide) throw new Error("No slide selected.");
  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;
    const s = shapes.addGeometricShape(geometry as any);
    s.left = opts.left ?? 100; s.top = opts.top ?? 100;
    s.width = opts.width ?? 150; s.height = opts.height ?? 100;
    (s as any).name = `${geometry}_${Date.now()}`;
    if (opts.fillColor) s.fill.setSolidColor(opts.fillColor);
    if (opts.lineColor) (s.lineFormat as any).color = opts.lineColor;
    if (opts.lineWeight !== undefined) s.lineFormat.weight = opts.lineWeight;
    await context.sync();
    s.load("id, name, type"); await context.sync();
    return s;
  });
}

export async function addTextBox(text: string, opts: ShapeOptions = {}): Promise<any> {
  const slide = await getSelectedSlide();
  if (!slide) throw new Error("No slide selected.");
  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;
    const tb = shapes.addTextBox(text);
    tb.left = opts.left ?? 80; tb.top = opts.top ?? 100;
    tb.width = opts.width ?? 600; tb.height = opts.height ?? 400;
    if (opts.fillColor) tb.fill.setSolidColor(opts.fillColor);
    await context.sync();
    // Apply font size if specified
    if (opts.fontSize !== undefined) {
      try {
        tb.textFrame.load("textRange"); await context.sync();
        tb.textFrame.textRange.font.size = opts.fontSize;
        await context.sync();
      } catch { /* font may not be available on this shape type */ }
    }
    return tb;
  });
}

export async function setShapeGeometry(sid: string, slid: string, geo: { left?: number; top?: number; width?: number; height?: number }): Promise<void> {
  return runPPT(async (context) => {
    const s = context.presentation.slides.getItem(slid).shapes.getItem(sid);
    if (geo.left !== undefined) s.left = geo.left;
    if (geo.top !== undefined) s.top = geo.top;
    if (geo.width !== undefined) s.width = geo.width;
    if (geo.height !== undefined) s.height = geo.height;
    await context.sync();
  });
}

export async function setShapeFill(sid: string, slid: string, color: string): Promise<void> {
  return runPPT(async (context) => {
    try {
      const shape = context.presentation.slides.getItem(slid).shapes.getItem(sid);
      shape.fill.setSolidColor(color);
      await context.sync();
    } catch {
      // Some shapes (pictures, groups, etc.) don't support fill — silently skip
    }
  });
}

export interface ParagraphSpec {
  text: string;
  fontSize?: number;
  bold?: boolean;
  fontColor?: string;
}

export async function addStructuredTextBox(
  paragraphs: ParagraphSpec[],
  opts: ShapeOptions = {}
): Promise<any> {
  const slide = await getSelectedSlide();
  if (!slide) throw new Error("No slide selected.");

  // Build full text first
  const fullText = paragraphs.map(p => p.text).join("\n");

  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;
    const tb = shapes.addTextBox(fullText);
    tb.left = opts.left ?? 80; tb.top = opts.top ?? 100;
    tb.width = opts.width ?? 600; tb.height = opts.height ?? 400;
    await context.sync();

    // Apply per-paragraph formatting
    try {
      const tr: any = tb.textFrame.textRange;
      tr.load("paragraphs/items/text, paragraphs/items/font/size, paragraphs/items/font/bold");
      await context.sync();

      const paraItems = tr.paragraphs.items;
      for (let i = 0; i < paragraphs.length && i < paraItems.length; i++) {
        const spec = paragraphs[i];
        if (spec.fontSize !== undefined) paraItems[i].font.size = spec.fontSize;
        if (spec.bold !== undefined) paraItems[i].font.bold = spec.bold;
        if (spec.fontColor) paraItems[i].font.color = spec.fontColor;
      }
      await context.sync();
    } catch {
      // Paragraph-level API not available — text is still created with default formatting
    }

    return tb;
  });
}

export interface TextOptions {
  fontSize?: number;
  bold?: boolean;
  fontName?: string;
  fontColor?: string;
}

export async function setShapeText(sid: string, slid: string, text: string, opts?: TextOptions): Promise<void> {
  return runPPT(async (context) => {
    const s = context.presentation.slides.getItem(slid).shapes.getItem(sid);
    s.textFrame.load("textRange"); await context.sync();
    s.textFrame.textRange.text = text; await context.sync();
    if (opts) {
      try {
        if (opts.fontSize !== undefined) s.textFrame.textRange.font.size = opts.fontSize;
        if (opts.bold !== undefined) s.textFrame.textRange.font.bold = opts.bold;
        if (opts.fontName) s.textFrame.textRange.font.name = opts.fontName;
        if (opts.fontColor) s.textFrame.textRange.font.color = opts.fontColor;
        await context.sync();
      } catch { /* font props may not be available */ }
    }
  });
}

export async function setShapeFontSize(sid: string, slid: string, fontSize: number): Promise<void> {
  return runPPT(async (context) => {
    const s = context.presentation.slides.getItem(slid).shapes.getItem(sid);
    try {
      s.textFrame.load("textRange"); await context.sync();
      s.textFrame.textRange.font.size = fontSize;
      await context.sync();
    } catch { /* shape might not have text */ }
  });
}

export async function deleteShape(sid: string, slid: string): Promise<void> {
  return runPPT(async (context) => {
    context.presentation.slides.getItem(slid).shapes.getItem(sid).delete();
    await context.sync();
  });
}

export async function applyStyleToAllShapes(slid: string, style: { fillColor?: string; lineColor?: string; lineWeight?: number; fontSize?: number }): Promise<number> {
  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slid).shapes;
    shapes.load("items/id"); await context.sync();
    for (const s of shapes.items) {
      if (style.fillColor) s.fill.setSolidColor(style.fillColor);
      if (style.lineColor) (s.lineFormat as any).color = style.lineColor;
      if (style.lineWeight !== undefined) s.lineFormat.weight = style.lineWeight;
      if (style.fontSize !== undefined) {
        try { s.textFrame.load("textRange/font"); await context.sync(); s.textFrame.textRange.font.size = style.fontSize; } catch { /* */ }
      }
      await context.sync();
    }
    return shapes.items.length;
  });
}

// ── Auto-Layout (detect overlaps, grid arrangement) ──────────────

interface ShapeInfo { id: string; name: string; left: number; top: number; width: number; height: number; }

/** Check if two rectangles overlap */
function rectsOverlap(a: ShapeInfo, b: ShapeInfo): boolean {
  return !(a.left + a.width <= b.left || b.left + b.width <= a.left ||
           a.top + a.height <= b.top || b.top + b.height <= a.top);
}

/** Detect and report overlapping shapes on a slide */
export async function detectOverlaps(slid: string): Promise<string[]> {
  const { getShapesOnSlide } = await import("./pptApi");
  const shapes = await getShapesOnSlide(slid);

  const infos: ShapeInfo[] = shapes.map(s => ({
    id: s.id, name: s.name || "unnamed",
    left: s.left, top: s.top, width: s.width, height: s.height,
  }));

  const overlaps: string[] = [];
  for (let i = 0; i < infos.length; i++) {
    for (let j = i + 1; j < infos.length; j++) {
      if (rectsOverlap(infos[i], infos[j])) {
        overlaps.push(`"${infos[i].name}" ↔ "${infos[j].name}"`);
      }
    }
  }
  return overlaps;
}

/** Auto-arrange shapes on a slide into a grid */
export async function autoLayoutShapes(slid: string, columns: number = 3): Promise<number> {
  const { getShapesOnSlide } = await import("./pptApi");
  const shapes = await getShapesOnSlide(slid);
  if (shapes.length === 0) return 0;

  const gap = 20;
  const margin = 40;
  const slideW = 960, slideH = 540;
  const availW = slideW - margin * 2 - gap * (columns - 1);
  const cellW = availW / columns;

  return runPPT(async (context) => {
    const slide = context.presentation.slides.getItem(slid);
    const pptShapes = slide.shapes;

    for (let i = 0; i < shapes.length; i++) {
      const col = i % columns;
      const row = Math.floor(i / columns);
      const cellH = Math.min(shapes[i].height || 80, (slideH - margin * 2) / Math.ceil(shapes.length / columns));
      const newLeft = margin + col * (cellW + gap);
      const newTop = margin + row * (cellH + gap);

      const s = pptShapes.getItem(shapes[i].id);
      s.left = newLeft;
      s.top = newTop;
      if (shapes[i].width && shapes[i].width > cellW) s.width = cellW - gap;
    }
    await context.sync();
    return shapes.length;
  });
}
