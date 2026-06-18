import { acquireToken } from "../auth/msal.js";
import type {
  AgentRequest,
  AgentResponse,
  BellaRawResponse,
  WorkdayDataResponse,
  WorkdayOptions,
} from "../types/index.js";

const WEBHOOK_URL = import.meta.env.VITE_N8N_WEBHOOK_URL as string;
const WORKDAY_DATA_URL = import.meta.env.VITE_N8N_WORKDAY_DATA_URL as string;
const WORKDAY_SUBMIT_URL = import.meta.env.VITE_N8N_WORKDAY_SUBMIT_URL as string;

export class AgentError extends Error {
  constructor(
    public readonly statusCode: number,
    message: string,
  ) {
    super(message);
    this.name = "AgentError";
  }
}

export async function callAgent(
  request: AgentRequest,
  webhookUrl = WEBHOOK_URL,
): Promise<AgentResponse> {
  console.log("[callAgent] webhookUrl:", webhookUrl);
  console.log("[callAgent] request:", JSON.stringify(request));

  const token = await acquireToken();
  console.log("[callAgent] token acquired, length:", token.length);

  const response = await fetch(webhookUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify(request),
  });

  console.log("[callAgent] response status:", response.status, response.statusText);
  console.log("[callAgent] response content-type:", response.headers?.get?.("content-type"));

  if (!response.ok) {
    const body = await response.text();
    console.error("[callAgent] error body:", body);
    throw new AgentError(response.status, body || response.statusText);
  }

  const text = await response.text();
  console.log("[callAgent] response body:", text);

  if (!text) {
    throw new AgentError(502, "Empty response from webhook");
  }

  let raw: BellaRawResponse;
  try {
    raw = JSON.parse(text) as BellaRawResponse;
    console.log("[callAgent] parsed raw keys:", Object.keys(raw));
  } catch {
    throw new AgentError(502, `Invalid JSON from webhook: ${text.slice(0, 200)}`);
  }

  return parseBellaResponse(raw);
}

// Fetches the Workday option lists for the current user. The n8n workflow
// resolves the user from the Bearer token, so the request carries no body.
export async function fetchWorkdayData(
  workdayDataUrl = WORKDAY_DATA_URL,
): Promise<WorkdayOptions> {
  const token = await acquireToken();

  const response = await fetch(workdayDataUrl, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    const body = await response.text();
    throw new AgentError(response.status, body || response.statusText);
  }

  const text = await response.text();
  if (!text) {
    throw new AgentError(502, "Empty response from Workday data endpoint");
  }

  let parsed: WorkdayDataResponse;
  try {
    parsed = JSON.parse(text) as WorkdayDataResponse;
  } catch {
    throw new AgentError(502, `Invalid JSON from Workday data endpoint: ${text.slice(0, 200)}`);
  }

  if (!parsed.success) {
    throw new AgentError(502, "Workday data endpoint returned success: false");
  }

  return parsed.data;
}

// Submits a requisition as multipart/form-data: the full form context as a
// single free-text `query` field, the binary documents (max 5) under
// `attachments`. Content-Type is left unset so the browser sets the multipart
// boundary. Returns the AI's response (HTML) to display on success.
export async function submitRequisition(
  query: string,
  attachments: File[],
  submitUrl = WORKDAY_SUBMIT_URL,
): Promise<string> {
  const token = await acquireToken();

  const form = new FormData();
  form.append("query", query);
  for (const file of attachments) {
    form.append("attachments", file, file.name);
  }

  const response = await fetch(submitUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    body: form,
  });

  if (!response.ok) {
    const body = await response.text();
    throw new AgentError(response.status, body || response.statusText);
  }

  const text = await response.text();
  if (!text) {
    throw new AgentError(502, "Empty response from webhook");
  }

  let raw: BellaRawResponse;
  try {
    raw = JSON.parse(text) as BellaRawResponse;
  } catch {
    throw new AgentError(502, `Invalid JSON from webhook: ${text.slice(0, 200)}`);
  }

  return raw.mail_template ?? "";
}

export function parseBellaResponse(raw: BellaRawResponse): AgentResponse {
  const template = raw.mail_template ?? "";

  const enquiriesMatch = template.match(/<b>Enquiries<\/b>:(.*?)(?=<b>Unanswered<\/b>|<b>Email<\/b>|$)/is);
  const unansweredMatch = template.match(/<b>Unanswered<\/b>:(.*?)(?=<b>Email<\/b>|$)/is);
  // Bella uses <b>Email</b>:<br><hr> as separator
  const emailMatch = template.match(/<b>Email<\/b>:\s*<br>\s*<hr>(.*?)$/is);

  return {
    enquiries: enquiriesMatch?.[1]?.trim() ?? "",
    unanswered: unansweredMatch?.[1]?.trim() ?? "",
    draft: emailMatch?.[1]?.trim() ?? template,
  };
}
