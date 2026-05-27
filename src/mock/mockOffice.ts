import { renderResult } from "../ui/draft.js";
import { clearError, showPanel } from "../ui/panels.js";
import { mockAgentResponse, mockMailItem } from "./mockData.js";

export function initializeMock(): void {
  showPanel("main-panel");

  document.getElementById("generate-btn")?.addEventListener("click", () => void handleMockGenerate());
  document.getElementById("insert-btn")?.addEventListener("click", () => handleMockInsert());
  document.getElementById("back-btn")?.addEventListener("click", () => showPanel("main-panel"));
}

async function handleMockGenerate(): Promise<void> {
  clearError();
  showPanel("loading");

  await new Promise((resolve) => setTimeout(resolve, 1200));

  console.info("[MOCK] Mail item:", mockMailItem);
  console.info("[MOCK] Agent response:", mockAgentResponse);

  renderResult(mockAgentResponse);
  showPanel("result");
}

function handleMockInsert(): void {
  console.info("[MOCK] Insert reply clicked — in real Outlook this would insert the draft.");
  showPanel("main-panel");
}
