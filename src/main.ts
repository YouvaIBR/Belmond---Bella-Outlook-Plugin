import { AgentError, callAgent } from "./api/agent.js";
import { initializeMock } from "./mock/mockOffice.js";
import { getMailItem, insertDraftReply } from "./services/office.js";
import { getDraftHtml, renderResult } from "./ui/draft.js";
import { clearError, showError, showPanel } from "./ui/panels.js";

const IS_MOCK = import.meta.env.VITE_MOCK === "true";

if (IS_MOCK) {
  initializeMock();
} else {
  Office.onReady(() => {
    void initializeAddin();
  });
}

async function initializeAddin(): Promise<void> {
  attachEventListeners();
  // Auto-trigger: generate immediately when the panel opens.
  // The generate button is hidden by default and shown only on error for retry.
  showPanel("loading");
  await handleGenerate();
}

function attachEventListeners(): void {
  document.getElementById("generate-btn")?.addEventListener("click", () => void handleGenerate());
  document.getElementById("insert-btn")?.addEventListener("click", () => void handleInsert());
  document.getElementById("back-btn")?.addEventListener("click", () => showPanel("main-panel"));
}

async function handleGenerate(): Promise<void> {
  clearError();
  showPanel("loading");

  const generateBtn = document.getElementById("generate-btn");
  if (generateBtn) generateBtn.hidden = true;

  try {
    const mailItem = await getMailItem();

    const response = await callAgent({
      emailBody: mailItem.body,
      emailSubject: mailItem.subject,
      emailFrom: mailItem.from,
      productCode: "LRS",
    });

    renderResult(response);
    showPanel("result");
  } catch (err) {
    if (generateBtn) generateBtn.hidden = false;

    if (err instanceof AgentError) {
      if (err.statusCode === 401 || err.statusCode === 403) {
        showPanel("access-denied");
        return;
      }
      showError(`Error ${err.statusCode.toString()}: ${err.message}`);
    } else if (err instanceof Error) {
      showError(err.message);
    } else {
      showError("An unexpected error occurred.");
    }
  }
}

async function handleInsert(): Promise<void> {
  const html = getDraftHtml();
  if (!html) {
    showError("No draft to insert.");
    showPanel("main-panel");
    return;
  }

  try {
    await insertDraftReply(html);
    showPanel("main-panel");
  } catch (err) {
    showPanel("result");
    showError(err instanceof Error ? err.message : "Failed to insert the draft reply.");
  }
}
