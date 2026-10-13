export const ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5];
export const MIN_ZOOM = ZOOM_STEPS[0] ?? 0.25;
export const MAX_ZOOM = ZOOM_STEPS[ZOOM_STEPS.length - 1] ?? 5;

export function clampZoom(zoom: number): number {
  return Number.isFinite(zoom) ? Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom)) : 1;
}

/** The next preset step above (direction 1) or below (-1) `zoom`, so a fitted zoom like 0.87 steps to 0.9 or 0.8. */
export function stepZoom(zoom: number, direction: 1 | -1): number {
  const epsilon = 1e-6;
  const next = direction > 0 ? ZOOM_STEPS.find((step) => step > zoom + epsilon) : [...ZOOM_STEPS].reverse().find((step) => step < zoom - epsilon);
  return next ?? clampZoom(zoom);
}

export function formatZoom(zoom: number): string {
  return `${Math.round(zoom * 100)}%`;
}
