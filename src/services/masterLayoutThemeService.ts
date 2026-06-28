/**
 * masterLayoutThemeService.ts — Slide Master, Layout, and Theme
 */

import { runPPT } from "./pptApi";

export async function getMasterDetails(): Promise<{ id: string; name: string; layoutCount: number }[]> {
  return runPPT(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id, items/name"); await context.sync();
    const r: any[] = [];
    for (const m of masters.items) {
      m.layouts.load("items/id"); await context.sync();
      r.push({ id: m.id, name: m.name, layoutCount: m.layouts.items.length });
    }
    return r;
  });
}

export async function applyLayoutToSlide(slideId: string, layoutId: string): Promise<void> {
  return runPPT(async (context) => {
    (context.presentation.slides.getItem(slideId) as any).layoutId = layoutId;
    await context.sync();
  });
}

export async function findLayoutByName(name: string): Promise<{ masterId: string; layoutId: string } | null> {
  return runPPT(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id, items/name"); await context.sync();
    for (const m of masters.items) {
      m.layouts.load("items/id, items/name"); await context.sync();
      for (const l of m.layouts.items) {
        if (l.name.toLowerCase().includes(name.toLowerCase())) {
          return { masterId: m.id, layoutId: l.id };
        }
      }
    }
    return null;
  });
}

export async function getAllLayouts(): Promise<{ masterId: string; masterName: string; layoutId: string; layoutName: string }[]> {
  return runPPT(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id, items/name"); await context.sync();
    const r: any[] = [];
    for (const m of masters.items) {
      m.layouts.load("items/id, items/name"); await context.sync();
      for (const l of m.layouts.items) {
        r.push({ masterId: m.id, masterName: m.name, layoutId: l.id, layoutName: l.name });
      }
    }
    return r;
  });
}

export async function setSlideBackground(slideId: string, color: string): Promise<void> {
  return runPPT(async (context) => {
    const slide = context.presentation.slides.getItem(slideId);
    // Try native background API first
    try {
      const fill: any = slide.background.fill;
      if (typeof fill.setSolidColor === "function") {
        fill.setSolidColor(color);
        await context.sync();
        return;
      }
      // Try direct color property
      fill.color = color;
      await context.sync();
      return;
    } catch {
      // Native API unavailable — fallback below
    }

    // Fallback: add a full-slide rectangle behind everything
    const shapes = slide.shapes;
    const rect = shapes.addGeometricShape("Rectangle" as any);
    rect.left = 0; rect.top = 0;
    rect.width = 960; // standard 16:9 slide width in points
    rect.height = 540; // standard 16:9 slide height
    rect.fill.setSolidColor(color);
    (rect as any).name = "_bg_rect";
    (rect as any).zOrder = "SendToBack";
    await context.sync();
  });
}

export async function getThemeDetails(): Promise<{ name: string }> {
  return runPPT(async (context) => {
    try {
      const pres: any = context.presentation;
      if (pres.theme) { pres.theme.load("name"); await context.sync(); return { name: pres.theme.name || "Default Theme" }; }
    } catch { /* */ }
    return { name: "Default Theme" };
  });
}

export async function addSlide(layoutId?: string): Promise<any> {
  return runPPT(async (context) => {
    const s: any = layoutId ? context.presentation.slides.add({ layoutId }) : context.presentation.slides.add();
    await context.sync();
    if (s.load) { s.load("id"); await context.sync(); }
    return s;
  });
}

export async function deleteSlide(slideId: string): Promise<void> {
  return runPPT(async (context) => {
    context.presentation.slides.getItem(slideId).delete();
    await context.sync();
  });
}
