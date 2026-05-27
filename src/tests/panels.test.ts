import { beforeEach, describe, expect, it } from "vitest";
import type { UIPanel } from "../types/index.js";
import { clearError, showError, showPanel } from "../ui/panels.js";

const ALL_PANELS: UIPanel[] = ["auth-loading", "access-denied", "main-panel", "loading", "result"];

function buildDom(): void {
  const sections = ALL_PANELS.map((id) => `<section id="${id}"></section>`).join("");
  document.body.innerHTML = `${sections}<div id="error-message"></div>`;
}

describe("showPanel", () => {
  beforeEach(buildDom);

  it("shows only the target panel", () => {
    showPanel("main-panel");
    expect(document.getElementById("main-panel")?.hidden).toBe(false);
    for (const id of ALL_PANELS) {
      if (id !== "main-panel") expect(document.getElementById(id)?.hidden).toBe(true);
    }
  });

  it("can switch between panels", () => {
    showPanel("loading");
    expect(document.getElementById("loading")?.hidden).toBe(false);
    showPanel("result");
    expect(document.getElementById("result")?.hidden).toBe(false);
    expect(document.getElementById("loading")?.hidden).toBe(true);
  });
});

describe("showError", () => {
  beforeEach(buildDom);

  it("sets error text and switches to main-panel", () => {
    showPanel("loading");
    showError("Something went wrong");
    expect(document.getElementById("error-message")?.textContent).toBe("Something went wrong");
    expect(document.getElementById("main-panel")?.hidden).toBe(false);
  });
});

describe("clearError", () => {
  beforeEach(buildDom);

  it("empties the error element", () => {
    const el = document.getElementById("error-message");
    if (el) el.textContent = "old error";
    clearError();
    expect(document.getElementById("error-message")?.textContent).toBe("");
  });
});

