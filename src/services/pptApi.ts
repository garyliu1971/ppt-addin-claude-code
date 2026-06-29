/**
 * pptApi.ts — Core PowerPoint API 1.10+ wrapper
 */

export interface LogEntry {
  level: "info" | "warn" | "error" | "success";
  message: string;
  timestamp: number;
}

export async function ensureOfficeReady(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.PowerPoint) resolve();
      else reject(new Error(`Unsupported host: ${info.host}`));
    });
  });
}

export async function runPPT<T>(
  fn: (context: PowerPoint.RequestContext) => Promise<T>
): Promise<T> {
  return PowerPoint.run(async (context) => fn(context));
}

export async function getSlides(): Promise<PowerPoint.Slide[]> {
  return runPPT(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    return slides.items;
  });
}

export async function getSelectedSlide(): Promise<PowerPoint.Slide | null> {
  return runPPT(async (context) => {
    const sel = context.presentation.getSelectedSlides();
    sel.load("items/id");
    await context.sync();
    return sel.items[0] || null;
  });
}

export async function getShapesOnSlide(slideId: string): Promise<PowerPoint.Shape[]> {
  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id, items/name, items/type, items/left, items/top, items/width, items/height");
    await context.sync();
    return shapes.items;
  });
}

export async function getSelectedShapes(): Promise<PowerPoint.Shape[]> {
  return runPPT(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id, items/name, items/type, items/left, items/top, items/width, items/height");
    await context.sync();
    return shapes.items;
  });
}

export async function getSlideMasters(): Promise<PowerPoint.SlideMaster[]> {
  return runPPT(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id, items/name");
    await context.sync();
    return masters.items;
  });
}

export async function getAllLayouts(): Promise<
  { masterId: string; masterName: string; layoutId: string; layoutName: string }[]
> {
  return runPPT(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id, items/name");
    await context.sync();
    const result: any[] = [];
    for (const master of masters.items) {
      const layouts = master.layouts;
      layouts.load("items/id, items/name");
      await context.sync();
      for (const layout of layouts.items) {
        result.push({ masterId: master.id, masterName: master.name, layoutId: layout.id, layoutName: layout.name });
      }
    }
    return result;
  });
}

export async function getPresentationInfo(): Promise<{ slideCount: number; title: string }> {
  return runPPT(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    const props: any = context.presentation.properties;
    props.load("title");
    await context.sync();
    return { slideCount: slides.items.length, title: props.title || "Untitled" };
  });
}

export async function getPageSetup(): Promise<{ width: number; height: number } | null> {
  return runPPT(async (context) => {
    // slideWidth/slideHeight in points (1 pt = 1/72 inch)
    const pres: any = context.presentation;
    pres.load("slideWidth, slideHeight");
    await context.sync();
    if (pres.slideWidth == null) return null;
    return { width: pres.slideWidth, height: pres.slideHeight };
  });
}
