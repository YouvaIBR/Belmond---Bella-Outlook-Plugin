import { acquireToken } from "../auth/msal.js";
import type { AgentRequest, AgentResponse, BellaRawResponse } from "../types/index.js";

const WEBHOOK_URL = import.meta.env.VITE_N8N_WEBHOOK_URL as string;

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
