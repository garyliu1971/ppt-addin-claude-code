/**
 * shapeService.ts — Shape creation and modification
 */

import { runPPT, getSelectedSlide } from "./pptApi";

export interface ShapeOptions {
  left?: number; top?: number; width?: number; height?: number;
  fillColor?: string; lineColor?: string; lineWeight?: number;
  fontSize?: number;
  /** 0.0 (opaque) to 1.0 (fully transparent) */
  transparency?: number;
  /** Show or hide the shape border */
  lineVisible?: boolean;
  /** Rotation in degrees */
  rotation?: number;
  /** Text margins (in points) */
  leftMargin?: number; rightMargin?: number; topMargin?: number; bottomMargin?: number;
  /** Vertical text alignment within the shape */
  verticalAlignment?: string;
  /** Horizontal text alignment */
  horizontalAlignment?: string;
  /** Text to place inside the shape (e.g. number badge, label) */
  text?: string;
  /** Color for the shape text */
  textColor?: string;
}

export async function addImage(imageUrl: string, opts: ShapeOptions = {}, slideId?: string): Promise<any> {
  let slide = slideId ? { id: slideId } as any : null;
  if (!slide) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slide = sel;
  }

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

export async function addShape(geometry: string, opts: ShapeOptions = {}, slideId?: string): Promise<any> {
  let slide = slideId ? { id: slideId } as any : null;
  if (!slide) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slide = sel;
  }
  console.log(`[addShape] slide=${slide.id}, geo="${geometry}", pos=(${opts.left ?? 100},${opts.top ?? 100})`);
  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;
    const s = shapes.addGeometricShape(geometry as any);
    s.left = opts.left ?? 100; s.top = opts.top ?? 100;
    s.width = opts.width ?? 150; s.height = opts.height ?? 100;
    (s as any).name = `${geometry}_${Date.now()}`;
    if (opts.fillColor) s.fill.setSolidColor(opts.fillColor);
    if (opts.transparency !== undefined) s.fill.transparency = opts.transparency;
    if (opts.lineColor) (s.lineFormat as any).color = opts.lineColor;
    if (opts.lineWeight !== undefined) s.lineFormat.weight = opts.lineWeight;
    if (opts.lineVisible !== undefined) s.lineFormat.visible = opts.lineVisible;
    if (opts.rotation !== undefined) (s as any).rotation = opts.rotation;
    await context.sync();
    // Copilot pattern: text inside shapes (numbered badges, labels in circles)
    if (opts.text) {
      try {
        s.textFrame.verticalAlignment = (opts.verticalAlignment as any) || "Middle";
        s.textFrame.textRange.text = opts.text;
        if (opts.fontSize) s.textFrame.textRange.font.size = opts.fontSize;
        if (opts.textColor) s.textFrame.textRange.font.color = opts.textColor;
        s.textFrame.textRange.font.bold = true;
        if (opts.horizontalAlignment) s.textFrame.textRange.paragraphFormat.horizontalAlignment = opts.horizontalAlignment as any;
        else s.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center" as any;
        await context.sync();
      } catch { /* text inside shape not supported */ }
    }
    s.load("id, name, type"); await context.sync();
    return s;
  });
}

export async function addTextBox(text: string, opts: ShapeOptions = {}, slideId?: string): Promise<any> {
  let slide = slideId ? { id: slideId } as any : null;
  if (!slide) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slide = sel;
  }
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
    // Apply text margins (Copilot: textFrame.leftMargin / rightMargin = 0)
    try {
      if (opts.leftMargin !== undefined) tb.textFrame.leftMargin = opts.leftMargin;
      if (opts.rightMargin !== undefined) tb.textFrame.rightMargin = opts.rightMargin;
      if (opts.topMargin !== undefined) tb.textFrame.topMargin = opts.topMargin;
      if (opts.bottomMargin !== undefined) tb.textFrame.bottomMargin = opts.bottomMargin;
      await context.sync();
    } catch { /* margins not supported on this shape */ }
    // Apply vertical alignment (Copilot: textFrame.verticalAlignment)
    try {
      if (opts.verticalAlignment) (tb.textFrame as any).verticalAlignment = opts.verticalAlignment;
      await context.sync();
    } catch { /* vertical alignment not supported */ }
    // Apply horizontal alignment (Copilot: paragraphFormat.horizontalAlignment)
    try {
      if (opts.horizontalAlignment) {
        const tr: any = tb.textFrame.textRange;
        tr.load("paragraphFormat"); await context.sync();
        tr.paragraphFormat.horizontalAlignment = opts.horizontalAlignment;
        await context.sync();
      }
    } catch { /* horizontal alignment not supported */ }
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
  /** Horizontal alignment for this paragraph: "Left" | "Center" | "Right" | "Justify" */
  alignment?: string;
  italic?: boolean;
}

export async function addStructuredTextBox(
  paragraphs: ParagraphSpec[],
  opts: ShapeOptions = {},
  slideId?: string
): Promise<any> {
  let slide = slideId ? { id: slideId } as any : null;
  if (!slide) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slide = sel;
  }

  if (!paragraphs || paragraphs.length === 0) throw new Error("No paragraphs provided");

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
      tr.load("paragraphs/items/text, paragraphs/items/font/size, paragraphs/items/font/bold, paragraphs/items/paragraphFormat");
      await context.sync();

      const paraItems = tr.paragraphs.items;
      for (let i = 0; i < paragraphs.length && i < paraItems.length; i++) {
        const spec = paragraphs[i];
        if (spec.fontSize !== undefined) paraItems[i].font.size = spec.fontSize;
        if (spec.bold !== undefined) paraItems[i].font.bold = spec.bold;
        if (spec.fontColor) paraItems[i].font.color = spec.fontColor;
        if (spec.italic !== undefined) paraItems[i].font.italic = spec.italic;
        // Copilot: paragraphFormat.horizontalAlignment per paragraph
        if (spec.alignment) {
          try {
            paraItems[i].paragraphFormat.horizontalAlignment = spec.alignment;
          } catch { /* alignment not available */ }
        }
      }
      await context.sync();
    } catch {
      // Paragraph-level API not available — text is still created with default formatting
    }

    // Apply text box-level formatting
    try {
      if (opts.verticalAlignment) (tb.textFrame as any).verticalAlignment = opts.verticalAlignment;
      if (opts.leftMargin !== undefined) tb.textFrame.leftMargin = opts.leftMargin;
      if (opts.rightMargin !== undefined) tb.textFrame.rightMargin = opts.rightMargin;
      if (opts.topMargin !== undefined) tb.textFrame.topMargin = opts.topMargin;
      if (opts.bottomMargin !== undefined) tb.textFrame.bottomMargin = opts.bottomMargin;
      await context.sync();
    } catch { /* margins/alignment not available */ }

    return tb;
  });
}

export interface TextOptions {
  fontSize?: number;
  bold?: boolean;
  fontName?: string;
  fontColor?: string;
  italic?: boolean;
  /** Horizontal text alignment: "Left" | "Center" | "Right" | "Justify" */
  alignment?: string;
  /** Vertical text alignment: "Top" | "Middle" | "Bottom" */
  verticalAlignment?: string;
  /** Text margins in points */
  leftMargin?: number; rightMargin?: number; topMargin?: number; bottomMargin?: number;
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
        if (opts.italic !== undefined) s.textFrame.textRange.font.italic = opts.italic;
        if (opts.alignment) (s.textFrame.textRange.paragraphFormat as any).horizontalAlignment = opts.alignment;
        if (opts.verticalAlignment) (s.textFrame as any).verticalAlignment = opts.verticalAlignment;
        if (opts.leftMargin !== undefined) s.textFrame.leftMargin = opts.leftMargin;
        if (opts.rightMargin !== undefined) s.textFrame.rightMargin = opts.rightMargin;
        if (opts.topMargin !== undefined) s.textFrame.topMargin = opts.topMargin;
        if (opts.bottomMargin !== undefined) s.textFrame.bottomMargin = opts.bottomMargin;
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

export async function applyStyleToAllShapes(slid: string, style: { fillColor?: string; lineColor?: string; lineWeight?: number; fontSize?: number; transparency?: number; lineVisible?: boolean }): Promise<number> {
  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slid).shapes;
    shapes.load("items/id"); await context.sync();
    for (const s of shapes.items) {
      if (style.fillColor) s.fill.setSolidColor(style.fillColor);
      if (style.transparency !== undefined) s.fill.transparency = style.transparency;
      if (style.lineColor) (s.lineFormat as any).color = style.lineColor;
      if (style.lineWeight !== undefined) s.lineFormat.weight = style.lineWeight;
      if (style.lineVisible !== undefined) s.lineFormat.visible = style.lineVisible;
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

// ── New Copilot-aligned operations ──────────────────────────────

/**
 * Insert an image directly from a base64 data URI.
 * Copilot: slide.shapes.addImage(base64, {left, top, width, height})
 */
export async function addImageFromBase64(base64: string, opts: ShapeOptions = {}, slideId?: string): Promise<any> {
  let slide = slideId ? { id: slideId } as any : null;
  if (!slide) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slide = sel;
  }

  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;
    // Copilot approach: try native addImage(base64) first (Office.js PictureAndCharts API)
    try {
      const img = (shapes as any).addImage(base64, { left: opts.left ?? 100, top: opts.top ?? 100, width: opts.width ?? 300, height: opts.height ?? 200 });
      await context.sync();
      if (opts.transparency !== undefined) {
        try { (img.fill as any).transparency = opts.transparency; await context.sync(); } catch { /* */ }
      }
      img.load("id, name, type"); await context.sync();
      return img;
    } catch {
      // native addImage not available — fall through
    }

    // Fallback: rectangle with picture fill
    const rect = shapes.addGeometricShape("Rectangle" as any);
    rect.left = opts.left ?? 100; rect.top = opts.top ?? 100;
    rect.width = opts.width ?? 300; rect.height = opts.height ?? 200;
    (rect as any).name = `Image_${Date.now()}`;

    try {
      await context.sync();
      const fill: any = rect.fill;
      fill.setImage(base64);
      if (opts.transparency !== undefined) fill.transparency = opts.transparency;
      await context.sync();
    } catch {
      // Picture fill failed — add a text label instead
      const label = shapes.addTextBox("[Image]");
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

/**
 * Apply formatting to any existing shape by slide ID + shape ID.
 * Supports all Copilot properties: fill, line, font, alignment, margins.
 */
export async function setShapeFormat(sid: string, slid: string, format: ShapeOptions & TextOptions): Promise<void> {
  return runPPT(async (context) => {
    const s = context.presentation.slides.getItem(slid).shapes.getItem(sid);

    // Geometry
    if (format.left !== undefined) s.left = format.left;
    if (format.top !== undefined) s.top = format.top;
    if (format.width !== undefined) s.width = format.width;
    if (format.height !== undefined) s.height = format.height;

    // Fill
    try {
      if (format.fillColor) s.fill.setSolidColor(format.fillColor);
      if (format.transparency !== undefined) s.fill.transparency = format.transparency;
    } catch { /* fill not supported */ }

    // Line format (Copilot: lineFormat.visible / color / weight)
    try {
      if (format.lineColor) (s.lineFormat as any).color = format.lineColor;
      if (format.lineWeight !== undefined) s.lineFormat.weight = format.lineWeight;
      if (format.lineVisible !== undefined) s.lineFormat.visible = format.lineVisible;
    } catch { /* line format not supported */ }

    // Rotation
    try {
      if (format.rotation !== undefined) (s as any).rotation = format.rotation;
    } catch { /* rotation not supported */ }

    // Text formatting
    try {
      s.textFrame.load("textRange"); await context.sync();
      if (format.fontSize !== undefined) s.textFrame.textRange.font.size = format.fontSize;
      if (format.bold !== undefined) s.textFrame.textRange.font.bold = format.bold;
      if (format.italic !== undefined) s.textFrame.textRange.font.italic = format.italic;
      if (format.fontName) s.textFrame.textRange.font.name = format.fontName;
      if (format.fontColor) s.textFrame.textRange.font.color = format.fontColor;
      // Copilot: paragraphFormat.horizontalAlignment
      if (format.alignment) (s.textFrame.textRange.paragraphFormat as any).horizontalAlignment = format.alignment;
    } catch { /* text formatting not available */ }

    // Text frame settings (Copilot: verticalAlignment, margins)
    try {
      if (format.verticalAlignment) (s.textFrame as any).verticalAlignment = format.verticalAlignment;
      if (format.leftMargin !== undefined) s.textFrame.leftMargin = format.leftMargin;
      if (format.rightMargin !== undefined) s.textFrame.rightMargin = format.rightMargin;
      if (format.topMargin !== undefined) s.textFrame.topMargin = format.topMargin;
      if (format.bottomMargin !== undefined) s.textFrame.bottomMargin = format.bottomMargin;
    } catch { /* text frame settings not available */ }

    await context.sync();
  });
}

// ── Compound card builder (Copilot: card + text in one call) ─────

export interface CardSpec {
  /** Card position */
  left: number; top: number; width: number; height: number;
  /** Card background color (default #0F2740) */
  fillColor?: string;
  /** Card border color (default #12314F) */
  lineColor?: string;
  /** Main heading text */
  heading: string;
  /** Heading font size (default 11) */
  headingSize?: number;
  /** Heading color (default #FFFFFF) */
  headingColor?: string;
  /** Subtitle / body text (optional) */
  subtitle?: string;
  /** Subtitle font size (default 9) */
  subtitleSize?: number;
  /** Subtitle color (default #9FB8CE) */
  subtitleColor?: string;
}

/**
 * Create a card WITH text in one atomic call.
 * AI can't forget to add text — card + heading + subtitle are created together.
 * @param spec Card specification
 * @param slideId Optional: target a specific slide instead of the selected one
 */
export async function addCard(spec: CardSpec, slideId?: string): Promise<any> {
  let slide = slideId ? { id: slideId } as any : null;
  if (!slide) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slide = sel;
  }
  console.log(`[addCard] targeting slide=${slide.id}, heading="${spec.heading}", pos=(${spec.left},${spec.top})`);

  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slide.id).shapes;

    // 1) Card background (RoundRectangle)
    const card = shapes.addGeometricShape("RoundRectangle" as any);
    card.left = spec.left; card.top = spec.top;
    card.width = spec.width; card.height = spec.height;
    card.fill.setSolidColor(spec.fillColor || "#0F2740");
    card.lineFormat.visible = true;
    (card.lineFormat as any).color = spec.lineColor || "#12314F";
    card.lineFormat.weight = 1;
    await context.sync();

    // 2) Heading text box INSIDE the card
    const headingBox = shapes.addTextBox(spec.heading);
    headingBox.left = spec.left + 10;
    headingBox.top = spec.top + 6;
    headingBox.width = spec.width - 20;
    headingBox.height = 20;
    await context.sync();
    headingBox.textFrame.textRange.font.size = spec.headingSize ?? 11;
    headingBox.textFrame.textRange.font.color = spec.headingColor || "#FFFFFF";
    headingBox.textFrame.textRange.font.bold = true;
    await context.sync();

    // 3) Subtitle text box (optional)
    if (spec.subtitle) {
      const subBox = shapes.addTextBox(spec.subtitle);
      subBox.left = spec.left + 10;
      subBox.top = spec.top + 28;
      subBox.width = spec.width - 20;
      subBox.height = 20;
      await context.sync();
      subBox.textFrame.textRange.font.size = spec.subtitleSize ?? 9;
      subBox.textFrame.textRange.font.color = spec.subtitleColor || "#9FB8CE";
      await context.sync();
    }

    card.load("id, name, type"); await context.sync();
    return card;
  });
}
