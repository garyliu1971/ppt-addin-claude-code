/**
 * eventService.ts — PowerPoint event handlers
 * Responds to shape selection changes and slide selection changes.
 * Uses PowerPoint API 1.10+ event binding.
 */

import { LogEntry } from "./pptApi";
import { getSelectedSlide, getSelectedShapes } from "./pptApi";

export type EventCallback = (entry: LogEntry) => void;
export type SlideChangeCallback = (slideId: string) => void;

/** Shape selection change handler */
let onShapeSelectionChangedHandler: (args?: any) => void;

/** Slide selection change handler */
let onSlideSelectionChangedHandler: (args?: any) => void;

/**
 * Register all PowerPoint event handlers.
 * @param onLog — callback to send log entries to the UI
 */
export async function registerEventHandlers(
  onLog: EventCallback,
  onSlideChange?: SlideChangeCallback
): Promise<void> {
  // ── Shape Selection Changed ─────────────────────────────────────
  onShapeSelectionChangedHandler = async () => {
    try {
      const shapes = await getSelectedShapes();
      if (shapes.length > 0) {
        const shapeNames = shapes.map((s) => s.name || s.id).join(", ");
        onLog({
          level: "info",
          message: `📐 Shape(s) selected: ${shapeNames} (${shapes.length} shape(s))`,
          timestamp: Date.now(),
        });
      } else {
        onLog({
          level: "info",
          message: "📐 Shape selection cleared",
          timestamp: Date.now(),
        });
      }
    } catch (err) {
      onLog({
        level: "error",
        message: `Shape selection handler error: ${err}`,
        timestamp: Date.now(),
      });
    }
  };

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onShapeSelectionChangedHandler,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        onLog({
          level: "success",
          message: "✅ Shape selection change listener registered",
          timestamp: Date.now(),
        });
      } else {
        onLog({
          level: "error",
          message: `❌ Failed to register shape selection listener: ${result.error.message}`,
          timestamp: Date.now(),
        });
      }
    }
  );

  // ── Slide Selection Changed ─────────────────────────────────────
  // We poll with a timer since Office JS PowerPoint doesn't have a
  // direct slide-selection-changed event in all versions.
  // Instead we use ActiveViewChanged or poll getSelectedSlides.

  onSlideSelectionChangedHandler = async () => {
    try {
      const slide = await getSelectedSlide();
      if (slide) {
        onLog({
          level: "info",
          message: `📑 Active slide changed: ID=${slide.id}`,
          timestamp: Date.now(),
        });
        // Update current slide in task pane
        if (onSlideChange) onSlideChange(slide.id);
      }
    } catch (err) {
      onLog({
        level: "error",
        message: `Slide change handler error: ${err}`,
        timestamp: Date.now(),
      });
    }
  };

  // Register ActiveViewChanged for slide navigation
  Office.context.document.addHandlerAsync(
    Office.EventType.ActiveViewChanged,
    onSlideSelectionChangedHandler,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        onLog({
          level: "success",
          message: "✅ Slide change listener registered",
          timestamp: Date.now(),
        });
      } else {
        onLog({
          level: "error",
          message: `❌ Failed to register slide change listener: ${result.error.message}`,
          timestamp: Date.now(),
        });
      }
    }
  );
}

/**
 * Unregister all event handlers (called on task pane close).
 */
export function unregisterEventHandlers(): void {
  if (onShapeSelectionChangedHandler) {
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      { handler: onShapeSelectionChangedHandler }
    );
  }
  if (onSlideSelectionChangedHandler) {
    Office.context.document.removeHandlerAsync(
      Office.EventType.ActiveViewChanged,
      { handler: onSlideSelectionChangedHandler }
    );
  }
}
