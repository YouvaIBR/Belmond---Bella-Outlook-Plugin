import type { UIPanel } from "../types/index.js";

const PANEL_IDS: UIPanel[] = ["auth-loading", "access-denied", "main-panel", "loading", "result"];

function getElement(id: string): HTMLElement {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Element #${id} not found`);
  return el;
}

export function showPanel(panel: UIPanel): void {
  for (const id of PANEL_IDS) {
    const el = document.getElementById(id);
    if (el) el.hidden = id !== panel;
  }
}

export function showError(message: string): void {
  const el = getElement("error-message");
  el.textContent = message;
  showPanel("main-panel");
}

export function clearError(): void {
  const el = document.getElementById("error-message");
  if (el) el.textContent = "";
}

