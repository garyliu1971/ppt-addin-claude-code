/**
 * chartTableService.ts — Chart & Table creation and modification
 * Tries native API first; falls back to shape-based visual representation.
 */

import { runPPT, getSelectedSlide, getShapesOnSlide } from "./pptApi";

export type ChartType = "ColumnClustered" | "ColumnStacked" | "BarClustered" | "BarStacked" | "Line" | "LineMarkers" | "Pie" | "Doughnut" | "Area" | "Scatter" | "Bubble";
export interface ChartData { categories: string[]; series: { name: string; values: number[] }[]; }

// ── Table ─────────────────────────────────────────────────────────

export async function addTable(
  data: { headers: string[]; rows: string[][] },
  opts: { left?: number; top?: number; width?: number; height?: number } = {},
  slideId?: string
): Promise<any> {
  let slideObj = slideId ? { id: slideId } as any : null;
  if (!slideObj) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slideObj = sel;
  }
  const left = opts.left ?? 50, top = opts.top ?? 120;
  const colW = (opts.width ?? 400) / data.headers.length;
  const rowH = 30, totalRows = data.rows.length + 1;

  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slideObj.id).shapes;

    // Try native addTable (Office 365 / API 1.7+)
    let nativeTable: any = null;
    try {
      const ts = (shapes as any).addTable(totalRows, data.headers.length, { hasHeaders: true });
      ts.left = left; ts.top = top;
      ts.width = opts.width ?? 400; ts.height = totalRows * rowH + 10;
      await context.sync();
      nativeTable = ts;
    } catch {
      // addTable API not available — fall through to visual table
    }

    if (nativeTable) {
      // Populate cells with error tolerance
      try {
        const tbl = (nativeTable as any).table;
        tbl.load("rows/items/cells/items"); await context.sync();
        for (let c = 0; c < data.headers.length; c++) {
          const cell = tbl.rows.items[0].cells.items[c];
          cell.textFrame.load("textRange"); await context.sync();
          cell.textFrame.textRange.text = data.headers[c];
          cell.textFrame.textRange.font.bold = true;
          cell.textFrame.textRange.font.size = 12;
        }
        for (let r = 0; r < data.rows.length; r++) {
          for (let c = 0; c < data.rows[r].length && c < data.headers.length; c++) {
            const cell = tbl.rows.items[r + 1].cells.items[c];
            cell.textFrame.load("textRange"); await context.sync();
            cell.textFrame.textRange.text = data.rows[r][c];
            cell.textFrame.textRange.font.size = 11;
          }
        }
        await context.sync();
        return nativeTable;
      } catch (e) {
        // Cell population failed but table exists — delete and fallback
        try { (nativeTable as any).delete(); await context.sync(); } catch { /* */ }
      }
    }

    // Fallback: visual table with text boxes — looks like a real table
    // Header row
    for (let c = 0; c < data.headers.length; c++) {
      const tb = shapes.addTextBox(data.headers[c]);
      tb.left = left + c * colW + 4; tb.top = top + 4;
      tb.width = colW - 8; tb.height = rowH - 8;
      tb.fill.setSolidColor("#1a1a2e");
      tb.textFrame.load("textRange/font"); await context.sync();
      tb.textFrame.textRange.font.color = "#FFFFFF";
      tb.textFrame.textRange.font.bold = true;
      tb.textFrame.textRange.font.size = 11;
    }
    // Data rows
    for (let r = 0; r < data.rows.length; r++) {
      const isAlt = r % 2 === 1;
      for (let c = 0; c < data.rows[r].length && c < data.headers.length; c++) {
        const tb = shapes.addTextBox(data.rows[r][c]);
        tb.left = left + c * colW + 4; tb.top = top + (r + 1) * rowH + 4;
        tb.width = colW - 8; tb.height = rowH - 8;
        if (isAlt) tb.fill.setSolidColor("#F0F0F5");
        else tb.fill.setSolidColor("#FFFFFF");
        tb.textFrame.load("textRange/font"); await context.sync();
        tb.textFrame.textRange.font.size = 10;
        tb.textFrame.textRange.font.color = "#333333";
      }
    }
    // Grid lines
    for (let r = 0; r <= totalRows; r++) {
      const hLine = shapes.addGeometricShape("Rectangle" as any);
      hLine.left = left; hLine.top = top + r * rowH;
      hLine.width = opts.width ?? 400; hLine.height = 1;
      hLine.fill.setSolidColor("#CCCCCC"); hLine.lineFormat.weight = 0;
    }
    for (let c = 0; c <= data.headers.length; c++) {
      const vLine = shapes.addGeometricShape("Rectangle" as any);
      vLine.left = left + c * colW; vLine.top = top;
      vLine.width = 1; vLine.height = totalRows * rowH;
      vLine.fill.setSolidColor("#CCCCCC"); vLine.lineFormat.weight = 0;
    }
    // Outer border
    const border = shapes.addGeometricShape("Rectangle" as any);
    border.left = left; border.top = top;
    border.width = opts.width ?? 400; border.height = totalRows * rowH;
    border.fill.clear(); border.lineFormat.color = "#999999"; border.lineFormat.weight = 1.5;
    await context.sync();
    return border;
  });
}

// ── Chart ─────────────────────────────────────────────────────────

export async function addChart(
  chartType: ChartType, data: ChartData,
  opts: { left?: number; top?: number; width?: number; height?: number; title?: string; hasLegend?: boolean } = {},
  slideId?: string
): Promise<any> {
  let slideObj = slideId ? { id: slideId } as any : null;
  if (!slideObj) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slideObj = sel;
  }
  const left = opts.left ?? 80, top = opts.top ?? 100;
  const w = opts.width ?? 500, h = opts.height ?? 350;

  return runPPT(async (context) => {
    const shapes = context.presentation.slides.getItem(slideObj.id).shapes;

    // Try native addChart
    try {
      const cs = (shapes as any).addChart(chartType, { left, top, width: w, height: h });
      await context.sync();
      const chart = (cs as any).chart;
      chart.load("title, legend"); await context.sync();
      if (opts.title) chart.title.text = opts.title;
      if (opts.hasLegend !== undefined) chart.legend.visible = opts.hasLegend;
      const hdr = [""].concat(data.series.map((s: any) => s.name));
      const rows = [hdr];
      for (let i = 0; i < data.categories.length; i++) {
        const row = [data.categories[i]];
        for (const s of data.series) row.push(String(s.values[i] ?? 0));
        rows.push(row);
      }
      (chart as any).setData(rows, "ByRows");
      await context.sync();
      console.log("[addChart] Native chart created successfully");
      return cs;
    } catch (e) {
      console.error("[addChart] Native API failed:", e);
      console.error("[addChart] Error details:", JSON.stringify(e, Object.getOwnPropertyNames(e)));
      /* fallback below */ }

    // Fallback: visual bar chart using rectangles
    const colors = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5"];
    const margin = { top: 50, bottom: 40, left: 50, right: 20 };
    const chartW = w - margin.left - margin.right;
    const chartH = h - margin.top - margin.bottom;

    const categories = data.categories;
    const allSeries = data.series;
    const maxVal = Math.max(...allSeries.flatMap((s: any) => s.values), 1);

    // Title
    if (opts.title) {
      const tb = shapes.addTextBox(opts.title);
      tb.left = left; tb.top = top - 8;
      tb.width = w; tb.height = 30;
      tb.textFrame.load("textRange/font"); await context.sync();
      tb.textFrame.textRange.font.size = 16;
      tb.textFrame.textRange.font.bold = true;
      tb.textFrame.textRange.font.color = "#333333";
    }

    const groupGap = chartW / categories.length;
    const barW = (groupGap * 0.7) / allSeries.length;

    for (let ci = 0; ci < categories.length; ci++) {
      const gx = left + margin.left + ci * groupGap;
      // Category label
      const lbl = shapes.addTextBox(categories[ci]);
      lbl.left = gx; lbl.top = top + margin.top + chartH + 2;
      lbl.width = groupGap; lbl.height = 20;
      lbl.textFrame.load("textRange/font"); await context.sync();
      lbl.textFrame.textRange.font.size = 9;
      lbl.textFrame.textRange.font.color = "#666666";

      for (let si = 0; si < allSeries.length; si++) {
        const val = allSeries[si].values[ci] ?? 0;
        const bh = Math.max((val / maxVal) * chartH, 2);
        const bx = gx + si * barW + groupGap * 0.15;
        const bar = shapes.addGeometricShape("Rectangle" as any);
        bar.left = bx; bar.top = top + margin.top + chartH - bh;
        bar.width = barW; bar.height = bh;
        bar.fill.setSolidColor(colors[si % colors.length]);
      }
    }

    // Legend
    if (opts.hasLegend !== false && allSeries.length > 1) {
      for (let si = 0; si < allSeries.length; si++) {
        const lx = left + w - 160 + si * 80;
        const dot = shapes.addGeometricShape("Rectangle" as any);
        dot.left = lx; dot.top = top + 4;
        dot.width = 10; dot.height = 10;
        dot.fill.setSolidColor(colors[si % colors.length]);
        const lt = shapes.addTextBox(allSeries[si].name);
        lt.left = lx + 14; lt.top = top + 2;
        lt.width = 66; lt.height = 14;
        lt.textFrame.load("textRange/font"); await context.sync();
        lt.textFrame.textRange.font.size = 9;
      }
    }

    // Border
    const border = shapes.addGeometricShape("Rectangle" as any);
    border.left = left; border.top = top;
    border.width = w; border.height = h;
    border.fill.clear(); border.lineFormat.weight = 0.5;
    await context.sync();
    return border;
  });
}

// ── Update Existing Table ────────────────────────────────────────

async function findTableShape(slideId?: string): Promise<PowerPoint.Shape | null> {
  if (!slideId) {
    const sel = await getSelectedSlide();
    if (!sel) return null;
    slideId = sel.id;
  }
  const shapes = await getShapesOnSlide(slideId);
  for (const s of shapes) {
    if ((s.type as string) === "Table" || (s.name || "").toLowerCase().includes("table")) return s;
  }
  return null;
}

export async function upsertTable(data: { headers: string[]; rows: string[][] }, slideId?: string): Promise<string> {
  let slideObj = slideId ? { id: slideId } as any : null;
  if (!slideObj) {
    const sel = await getSelectedSlide();
    if (!sel) throw new Error("No slide selected.");
    slideObj = sel;
  }
  const existing = await findTableShape(slideObj.id);

  if (existing) {
    return runPPT(async (context) => {
      const shape = context.presentation.slides.getItem(slideObj.id).shapes.getItem(existing.id);
      try {
        const tbl: any = (shape as any).table;
        tbl.load("rows/items/cells/items"); await context.sync();
        for (let c = 0; c < data.headers.length && c < tbl.rows.items[0]?.cells.items.length; c++) {
          const cell = tbl.rows.items[0].cells.items[c];
          cell.textFrame.load("textRange"); await context.sync();
          cell.textFrame.textRange.text = data.headers[c];
        }
        for (let r = 0; r < data.rows.length && r + 1 < tbl.rows.items.length; r++) {
          for (let c = 0; c < data.rows[r].length && c < tbl.rows.items[r + 1]?.cells.items.length; c++) {
            const cell = tbl.rows.items[r + 1].cells.items[c];
            cell.textFrame.load("textRange"); await context.sync();
            cell.textFrame.textRange.text = data.rows[r][c];
          }
        }
        await context.sync();
        return `Updated existing table with ${data.rows.length} row(s)`;
      } catch { shape.delete(); await context.sync(); await addTable(data, {}, slideObj.id); return "Replaced existing table"; }
    });
  }
  await addTable(data, {}, slideObj.id);
  return `Created new table with ${data.rows.length} row(s)`;
}
