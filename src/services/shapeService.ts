/**
 * shapeService.ts — Shape creation and modification
 */

import { runPPT, getSelectedSlide } from "./pptApi";

export interface ShapeOptions {
  left?: number; top?: number; width?: number; height?: number;
  fillColor?: string; lineColor?: string; lineWeight?: number;
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
    tb.width = opts.width ?? 300; tb.height = opts.height ?? 60;
    if (opts.fillColor) tb.fill.setSolidColor(opts.fillColor);
    await context.sync();
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

export async function setShapeText(sid: string, slid: string, text: string): Promise<void> {
  return runPPT(async (context) => {
    const s = context.presentation.slides.getItem(slid).shapes.getItem(sid);
    s.textFrame.load("textRange"); await context.sync();
    s.textFrame.textRange.text = text; await context.sync();
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
