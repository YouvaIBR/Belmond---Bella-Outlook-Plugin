import { AgentError, fetchWorkdayData, submitRequisition } from "./api/agent.js";
import { initializeWorkdayMock } from "./mock/mockOffice.js";
import { getPdfAttachments } from "./services/attachments.js";
import { getMailItem, getUserDisplayName } from "./services/office.js";
import type { WorkdayOption, WorkdayOptions, WorkdayPanel } from "./types/index.js";
import { sanitizeHtml } from "./ui/draft.js";

const IS_MOCK = import.meta.env.VITE_MOCK === "true";

const WORKDAY_PANEL_IDS: WorkdayPanel[] = [
  "auth-loading",
  "access-denied",
  "workday-actions",
  "workday-form",
  "loading",
  "result",
];

// The selects that map directly to a single Workday option list.
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

let options: WorkdayOptions | null = null;
// PDF attachments pulled from the email (empty → user must pick from computer).
let emailAttachments: File[] = [];

function showPanel(panel: WorkdayPanel): void {
  for (const id of WORKDAY_PANEL_IDS) {
    const el = document.getElementById(id);
    if (el) el.hidden = id !== panel;
  }
}

function showError(message: string): void {
  const el = document.getElementById("form-error");
  if (el) {
    el.textContent = message;
    el.hidden = false;
  }
}

function clearError(): void {
  const el = document.getElementById("form-error");
  if (el) {
    el.textContent = "";
    el.hidden = true;
  }
}

if (IS_MOCK) {
  initializeWorkdayMock();
} else {
  Office.onReady(() => {
    void initializeWorkdayAddin();
  });
}

async function initializeWorkdayAddin(): Promise<void> {
  attachEventListeners();
  renderUserName();
  showPanel("auth-loading");

  try {
    options = await fetchWorkdayData();
  } catch (err) {
    handleError(err);
    return;
  }

  showPanel("workday-actions");
}

function renderUserName(): void {
  const name = getUserDisplayName();
  const el = document.getElementById("user-name");
  if (el && name) {
    el.textContent = `Signed in as ${name}`;
    el.hidden = false;
  }
}

function attachEventListeners(): void {
  document
    .querySelector('[data-action="submit-requisition"]')
    ?.addEventListener("click", () => void openRequisitionForm());
  document
    .getElementById("form-back-btn")
    ?.addEventListener("click", () => showPanel("workday-actions"));
  document
    .getElementById("result-done-btn")
    ?.addEventListener("click", () => showPanel("workday-actions"));

  const form = document.getElementById("requisition-form");
  form?.addEventListener("submit", (e) => {
    e.preventDefault();
    void handleSubmit();
  });
  // Any field change re-evaluates whether submit can be enabled.
  form?.addEventListener("input", () => syncSubmitEnabled());
  form?.addEventListener("change", () => syncSubmitEnabled());
}

// ── Form setup ────────────────────────────────────────────────────────────

async function openRequisitionForm(): Promise<void> {
  clearError();
  if (!options) return;

  populateSelects(options);

  try {
    emailAttachments = await getPdfAttachments();
  } catch {
    emailAttachments = [];
  }
  renderAttachments();

  syncSubmitEnabled();
  showPanel("workday-form");
}

function populateSelects(opts: WorkdayOptions): void {
  for (const { id, key } of SELECT_FIELDS) {
    const select = document.getElementById(id) as HTMLSelectElement | null;
    if (!select) continue;
    select.replaceChildren();

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Select…";
    select.append(placeholder);

    for (const opt of opts[key]) {
      const option = document.createElement("option");
      option.value = opt.id;
      option.textContent = opt.descriptor;
      select.append(option);
    }
  }
}

// Shows the email PDFs as a read-only list, or reveals the file picker if none.
function renderAttachments(): void {
  const list = document.getElementById("attachment-list") as HTMLElement;
  const picker = document.getElementById("attachment-picker") as HTMLElement;
  list.replaceChildren();

  if (emailAttachments.length > 0) {
    picker.hidden = true;
    for (const file of emailAttachments) {
      const item = document.createElement("div");
      item.className = "wd-attachment-item";
      item.textContent = file.name;
      list.append(item);
    }
  } else {
    picker.hidden = false;
  }
}

function getPickedFiles(): File[] {
  const input = document.getElementById("f-fileInput") as HTMLInputElement | null;
  return input?.files ? Array.from(input.files) : [];
}

function getAttachments(): File[] {
  return emailAttachments.length > 0 ? emailAttachments : getPickedFiles();
}

// All form fields are optional; the only requirement is at least one attachment
// (from the email, or picked from the computer when the email has none).
function syncSubmitEnabled(): void {
  const btn = document.getElementById("submit-requisition-btn") as HTMLButtonElement;
  btn.disabled = getAttachments().length === 0;
}

// ── Submit ──────────────────────────────────────────────────────────────────

// Returns the chosen option, or null when the field is left empty (so the
// backend AI step can infer it from the email).
function selectedOption(id: string, key: keyof WorkdayOptions): WorkdayOption | null {
  const select = document.getElementById(id) as HTMLSelectElement;
  const value = select.value;
  if (!value) return null;
  return options?.[key].find((o) => o.id === value) ?? null;
}

function val(id: string): string {
  return (document.getElementById(id) as HTMLInputElement | HTMLTextAreaElement).value.trim();
}

// Formats one Workday field for the query: includes both the selected id and its
// label so the backend can use the id directly (no re-matching needed).
function fieldLine(label: string, opt: WorkdayOption | null): string {
  if (!opt) return `${label}: (not specified)`;
  return `${label}: ${opt.descriptor} [id: ${opt.id}]`;
}

// Builds the single free-text `query` the backend agent expects: every form
// value the user filled (ids + labels + free text) plus the email content.
// Empty fields are sent as "(not specified)" so the backend can infer them.
function buildQuery(mailItem: { subject: string; from: string; body: string }): string {
  const highPriority = (document.getElementById("f-highPriority") as HTMLInputElement).checked;

  const lines = [
    "Submit a Workday purchase requisition with the following details.",
    "",
    "## Requisition",
    fieldLine("Requisition type", selectedOption("f-requisitionType", "requisitionType")),
    fieldLine("Company", selectedOption("f-company", "company")),
    fieldLine("Currency", selectedOption("f-currency", "currency")),
    fieldLine("Requester", selectedOption("f-requester", "requester")),
    `High priority: ${highPriority ? "yes" : "no"}`,
    "",
    "## Line item",
    `Item description: ${val("f-itemDescription") || "(not specified)"}`,
    fieldLine("Business unit", selectedOption("f-businessUnit", "businessUnit")),
    fieldLine("Cost center", selectedOption("f-costCenter", "costCenter")),
    fieldLine("Spend category", selectedOption("f-spendCategory", "spendCategory")),
    fieldLine("Supplier", selectedOption("f-supplier", "supplier")),
    fieldLine("Ship-to contact", selectedOption("f-shipToContact", "shipToContact")),
    `Quantity: ${val("f-quantity") || "(not specified)"}`,
    fieldLine("Unit of measure", selectedOption("f-unitOfMeasure", "unitOfMeasure")),
    `Unit cost: ${val("f-unitCost") || "(not specified)"}`,
    `Supplier item identifier: ${val("f-supplierItemIdentifier") || "(not specified)"}`,
    `Memo: ${val("f-memo") || "(not specified)"}`,
    "",
    "## Source email",
    `From: ${mailItem.from}`,
    `Subject: ${mailItem.subject}`,
    "Body:",
    mailItem.body,
  ];

  return lines.join("\n");
}

async function handleSubmit(): Promise<void> {
  clearError();

  const attachments = getAttachments();
  if (attachments.length === 0) {
    showError("At least one attachment is required.");
    return;
  }
  if (attachments.length > 5) {
    showError("You can attach at most 5 documents.");
    return;
  }

  showPanel("loading");

  try {
    const mailItem = await getMailItem();
    const response = await submitRequisition(buildQuery(mailItem), attachments);

    renderResultContent(response);
    showPanel("result");
  } catch (err) {
    showPanel("workday-form");
    handleError(err);
  }
}

function renderResultContent(html: string): void {
  const el = document.getElementById("result-content");
  if (el) el.innerHTML = sanitizeHtml(html); // nosec: sanitized
}

function handleError(err: unknown): void {
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
