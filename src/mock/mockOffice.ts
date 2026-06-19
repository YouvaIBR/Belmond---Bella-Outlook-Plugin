import type { WorkdayOptions, WorkdayPanel } from "../types/index.js";
import { renderResult, sanitizeHtml } from "../ui/draft.js";
import { clearError, showPanel } from "../ui/panels.js";
import {
  mockAgentResponse,
  mockMailItem,
  mockRequisitionResponse,
  mockUserName,
  mockWorkdayOptions,
} from "./mockData.js";

function renderMockUserName(): void {
  const el = document.getElementById("user-name");
  if (el) {
    el.textContent = `Signed in as ${mockUserName}`;
    el.hidden = false;
  }
}

export function initializeMock(): void {
  showPanel("main-panel");
  renderMockUserName();

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

// ── Workday actions / requisition form flow ─────────────────────────────────

const WORKDAY_PANEL_IDS: WorkdayPanel[] = [
  "auth-loading",
  "access-denied",
  "workday-actions",
  "workday-form",
  "loading",
  "result",
];

function showWorkdayPanel(panel: WorkdayPanel): void {
  for (const id of WORKDAY_PANEL_IDS) {
    const el = document.getElementById(id);
    if (el) el.hidden = id !== panel;
  }
}

const SELECT_FIELDS: { id: string; key: keyof WorkdayOptions }[] = [
  { id: "f-requisitionType", key: "requisitionType" },
  { id: "f-company", key: "company" },
  { id: "f-currency", key: "currency" },
  { id: "f-requester", key: "requester" },
  { id: "f-businessUnit", key: "businessUnit" },
  { id: "f-costCenter", key: "costCenter" },
  { id: "f-spendCategory", key: "spendCategory" },
  { id: "f-supplier", key: "supplier" },
  { id: "f-shipToContact", key: "shipToContact" },
  { id: "f-unitOfMeasure", key: "unitOfMeasure" },
];

export function initializeWorkdayMock(): void {
  showWorkdayPanel("auth-loading");
  renderMockUserName();

  document
    .querySelector('[data-action="submit-requisition"]')
    ?.addEventListener("click", () => openMockForm());
  document
    .getElementById("form-back-btn")
    ?.addEventListener("click", () => showWorkdayPanel("workday-actions"));
  document
    .getElementById("result-done-btn")
    ?.addEventListener("click", () => showWorkdayPanel("workday-actions"));

  const form = document.getElementById("requisition-form") as HTMLFormElement;
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    void handleMockSubmit();
  });
  form.addEventListener("input", () => syncMockSubmit());
  form.addEventListener("change", () => syncMockSubmit());

  // Simulate the data fetch on open.
  setTimeout(() => showWorkdayPanel("workday-actions"), 500);
}

function openMockForm(): void {
  for (const { id, key } of SELECT_FIELDS) {
    const select = document.getElementById(id) as HTMLSelectElement;
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Select…";
    select.append(placeholder);
    for (const opt of mockWorkdayOptions[key]) {
      const option = document.createElement("option");
      option.value = opt.id;
      option.textContent = opt.descriptor;
      select.append(option);
    }
  }

  // No real email attachments in mock → reveal the file picker.
  (document.getElementById("attachment-list") as HTMLElement).replaceChildren();
  (document.getElementById("attachment-picker") as HTMLElement).hidden = false;

  // Mock prefill: re-run each time the user adds files (email has none here).
  const fileInput = document.getElementById("f-fileInput") as HTMLInputElement;
  fileInput.addEventListener("change", () => {
    if ((fileInput.files?.length ?? 0) > 0) void runMockPrefill();
  });

  syncMockSubmit();
  showWorkdayPanel("workday-form");

  // Always prefill on open (email text alone in mock).
  void runMockPrefill();
}

async function runMockPrefill(): Promise<void> {
  const hint = document.getElementById("form-prefilling");
  if (hint) hint.hidden = false;
  await new Promise((resolve) => setTimeout(resolve, 1000));
  if (hint) hint.hidden = true;

  // Fill the first option of each select + sample text values.
  for (const { id, key } of SELECT_FIELDS) {
    const select = document.getElementById(id) as HTMLSelectElement;
    const first = mockWorkdayOptions[key][0];
    if (first) select.value = first.id;
  }
  (document.getElementById("f-itemDescription") as HTMLInputElement).value =
    "Architecture consultancy — pool area redesign";
  (document.getElementById("f-unitCost") as HTMLInputElement).value = "1000";
  (document.getElementById("f-quantity") as HTMLInputElement).value = "1";
  (document.getElementById("f-supplierItemIdentifier") as HTMLInputElement).value = "TEST1234";
  console.info("[MOCK] Prefill applied from email + documents.");
  syncMockSubmit();
}

function syncMockSubmit(): void {
  const btn = document.getElementById("submit-requisition-btn") as HTMLButtonElement;
  const fileInput = document.getElementById("f-fileInput") as HTMLInputElement;
  btn.disabled = (fileInput.files?.length ?? 0) === 0;
}

async function handleMockSubmit(): Promise<void> {
  showWorkdayPanel("loading");
  console.info("[MOCK] Submit requisition — mail item:", mockMailItem);
  await new Promise((resolve) => setTimeout(resolve, 1200));

  const el = document.getElementById("result-content");
  if (el) el.innerHTML = sanitizeHtml(mockRequisitionResponse); // nosec: sanitized
  showWorkdayPanel("result");
}
