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

// ── Theme & Design ────────────────────────────────────────────────

/** Known built-in PowerPoint theme names */
const BUILTIN_THEMES = [
  "Office", "Facet", "Ion", "Ion Boardroom", "Organic", "Slice",
  "Wisp", "Integral", "Retrospect", "Parallax", "Celestial", "Vapor Trail",
  "Badge", "Gallery", "Main Event", "Mesh", "Savon", "View",
  "Berlin", "Crop", "Circuit", "Depth", "Droplet", "Headlines",
  "Madison", "Mesh", "Metropolitan", "Parcel", "Quotable", "Wood Type",
];

export function listAvailableThemes(): string[] {
  return BUILTIN_THEMES;
}

/** Apply a theme by name to the presentation */
export async function applyTheme(themeName: string): Promise<string> {
  return runPPT(async (context) => {
    // Approach 1: Try native theme API (PowerPoint 365+)
    try {
      const pres: any = context.presentation;
      // Check if applyTheme method exists (newer API)
      if (typeof pres.applyTheme === "function") {
        pres.applyTheme(themeName);
        await context.sync();
        return `Applied theme "${themeName}"`;
      }
      // Check if theme.name is writable
      if (pres.theme) {
        pres.theme.name = themeName;
        await context.sync();
        // Verify it actually changed
        pres.theme.load("name"); await context.sync();
        if (pres.theme.name === themeName) {
          return `Applied theme "${themeName}"`;
        }
      }
    } catch (e) {
      console.warn("[applyTheme] Native API failed:", e);
    }

    // Approach 2: Fallback — apply a matching design scheme
    const themeMap: Record<string, string> = {
      "slice": "slate gray",
      "office": "clean white",
      "facet": "ocean blue",
      "ion": "modern dark",
      "organic": "forest green",
      "wisp": "sunset orange",
      "retrospect": "rose gold",
      "celestial": "ocean blue",
      "vapor trail": "slate gray",
      "berlin": "modern dark",
      "circuit": "ocean blue",
      "depth": "modern dark",
      "droplet": "ocean blue",
      "headlines": "slate gray",
      "metropolitan": "modern dark",
      "wood type": "forest green",
    };
    const schemeName = themeMap[themeName.toLowerCase()];
    if (schemeName) {
      // Apply to all slides
      const slides = context.presentation.slides;
      slides.load("items/id"); await context.sync();
      const scheme = COLOR_SCHEMES[schemeName];
      if (scheme) {
        for (const slide of slides.items) {
          try {
            (slide.background.fill as any).setSolidColor(scheme.bg);
          } catch {
            const shapes = slide.shapes;
            const rect = shapes.addGeometricShape("Rectangle" as any);
            rect.left = 0; rect.top = 0;
            rect.width = 960; rect.height = 540;
            rect.fill.setSolidColor(scheme.bg);
            (rect as any).zOrder = "SendToBack";
          }
        }
        await context.sync();
        return `Applied theme "${themeName}" via "${schemeName}" design scheme`;
      }
    }

    return `Theme "${themeName}" not available — applied best-effort styling`;
  });
}

/** Predefined color schemes for slide design */
const COLOR_SCHEMES: Record<string, { bg: string; accent: string; text: string }> = {
  "modern dark": { bg: "#1a1a2e", accent: "#667eea", text: "#e2e8f0" },
  "ocean blue": { bg: "#0f172a", accent: "#38bdf8", text: "#f0f9ff" },
  "forest green": { bg: "#14532d", accent: "#4ade80", text: "#f0fdf4" },
  "sunset orange": { bg: "#2d1b00", accent: "#fb923c", text: "#fff7ed" },
  "rose gold": { bg: "#1a0a0a", accent: "#f43f5e", text: "#fff1f2" },
  "clean white": { bg: "#ffffff", accent: "#6366f1", text: "#1e293b" },
  "slate gray": { bg: "#1e293b", accent: "#94a3b8", text: "#f8fafc" },
};

export function listDesignSchemes(): string[] {
  return Object.keys(COLOR_SCHEMES);
}

/** Apply a designer color scheme to the current slide */
export async function applyDesignScheme(slideId: string, schemeName: string): Promise<string> {
  const scheme = COLOR_SCHEMES[schemeName.toLowerCase()];
  if (!scheme) throw new Error(`Unknown scheme "${schemeName}". Available: ${Object.keys(COLOR_SCHEMES).join(", ")}`);

  return runPPT(async (context) => {
    // Set background
    const slide = context.presentation.slides.getItem(slideId);
    try {
      (slide.background.fill as any).setSolidColor(scheme.bg);
    } catch {
      const shapes = slide.shapes;
      const rect = shapes.addGeometricShape("Rectangle" as any);
      rect.left = 0; rect.top = 0;
      rect.width = 960; rect.height = 540;
      rect.fill.setSolidColor(scheme.bg);
      (rect as any).zOrder = "SendToBack";
    }
    await context.sync();

    // Apply accent to all shapes
    const shapes = slide.shapes;
    shapes.load("items/id, items/type"); await context.sync();
    for (const s of shapes.items) {
      try { s.fill.setSolidColor(scheme.accent); } catch { /* */ }
      try {
        s.textFrame.load("textRange/font"); await context.sync();
        s.textFrame.textRange.font.color = scheme.text;
      } catch { /* */ }
    }
    await context.sync();
    return `Applied "${schemeName}" design scheme`;
  });
}

export async function addSlide(layoutId?: string): Promise<any> {
  return runPPT(async (context) => {
    const slides = context.presentation.slides;
    if (layoutId) {
      slides.add({ layoutId } as any);
    } else {
      slides.add();
    }
    await context.sync();
    // Reload and get the last slide
    slides.load("items/id"); await context.sync();
    return slides.items[slides.items.length - 1];
  });
}

export async function deleteSlide(slideId: string): Promise<void> {
  return runPPT(async (context) => {
    context.presentation.slides.getItem(slideId).delete();
    await context.sync();
  });
}

// ── New Slide-Level Operations ────────────────────────────────────

/** Delete a slide by its 1-based index */
export async function deleteSlideByIndex(index: number): Promise<string> {
  return runPPT(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id"); await context.sync();
    if (index < 1 || index > slides.items.length) {
      throw new Error(`Slide index ${index} out of range (1-${slides.items.length}).`);
    }
    const target = slides.items[index - 1];
    target.delete();
    await context.sync();
    return `Deleted slide ${index}`;
  });
}

/** Get all slides with their indices */
export async function getSlidesWithIndex(): Promise<{ index: number; id: string }[]> {
  return runPPT(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id"); await context.sync();
    return slides.items.map((s, i) => ({ index: i + 1, id: s.id }));
  });
}

/** Set or change the title of a slide */
export async function setSlideTitle(slideId: string, title: string): Promise<void> {
  return runPPT(async (context) => {
    const slide = context.presentation.slides.getItem(slideId);
    const shapes = slide.shapes;
    shapes.load("items/id, items/name, items/type"); await context.sync();

    // Look for a title placeholder or text box
    let titleShape: PowerPoint.Shape | null = null;
    for (const s of shapes.items) {
      const name = (s.name || "").toLowerCase();
      const type = (s as any).type || "";
      // Title placeholder or shape named "Title"
      if (name.includes("title") || type === "Title") {
        titleShape = s;
        break;
      }
    }

    if (titleShape) {
      // Update existing title
      titleShape.textFrame.load("textRange"); await context.sync();
      titleShape.textFrame.textRange.text = title;
      titleShape.textFrame.textRange.font.size = 36;
      titleShape.textFrame.textRange.font.bold = true;
      await context.sync();
    } else {
      // No title found — add a text box at the top
      const tb = shapes.addTextBox(title);
      tb.left = 80; tb.top = 40;
      tb.width = 800; tb.height = 80;
      (tb as any).name = "Title";
      tb.textFrame.load("textRange/font"); await context.sync();
      tb.textFrame.textRange.font.size = 36;
      tb.textFrame.textRange.font.bold = true;
      tb.textFrame.textRange.font.color = "#1a1a2e";
      await context.sync();
    }
  });
}

/** Move a slide to a new 1-based index */
export async function moveSlide(slideId: string, toIndex: number): Promise<void> {
  return runPPT(async (context) => {
    const slide = context.presentation.slides.getItem(slideId);
    (slide as any).moveTo(toIndex - 1);
    await context.sync();
  });
}

/** Duplicate a slide */
export async function duplicateSlide(slideId: string): Promise<any> {
  return runPPT(async (context) => {
    const slide = context.presentation.slides.getItem(slideId);
    const dup = (slide as any).duplicate();
    await context.sync();
    if (dup && dup.load) { dup.load("id"); await context.sync(); }
    return dup;
  });
}

/** Add a slide with a title */
export async function addSlideWithTitle(title: string, layoutId?: string): Promise<any> {
  return runPPT(async (context) => {
    const slides = context.presentation.slides;

    // slides.add() returns void in most API versions — we need to
    // reload the collection and grab the last slide
    if (layoutId) {
      slides.add({ layoutId } as any);
    } else {
      slides.add();
    }
    await context.sync();

    // Reload all slides and get the last one (newly added)
    slides.load("items/id"); await context.sync();
    const s = slides.items[slides.items.length - 1];
    if (!s) throw new Error("Failed to create slide.");

    // Try to set the title on the new slide
    const shapes = s.shapes;
    shapes.load("items/id, items/name, items/type"); await context.sync();

    let titleShape: any = null;
    for (const sh of shapes.items) {
      const name = (sh.name || "").toLowerCase();
      if (name.includes("title")) { titleShape = sh; break; }
    }

    if (titleShape) {
      titleShape.textFrame.load("textRange"); await context.sync();
      titleShape.textFrame.textRange.text = title;
      titleShape.textFrame.textRange.font.size = 36;
      titleShape.textFrame.textRange.font.bold = true;
    } else {
      const tb = s.shapes.addTextBox(title);
      (tb as any).name = "Title";
      tb.left = 80; tb.top = 40;
      tb.width = 800; tb.height = 80;
      await context.sync();
    }

    await context.sync();
    return s;
  });
}
