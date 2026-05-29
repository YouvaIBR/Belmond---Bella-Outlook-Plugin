export interface MailItem {
  subject: string;
  from: string;
  body: string;
}

export interface AgentRequest {
  emailBody: string;
  emailSubject: string;
  emailFrom: string;
  productCode: string;
}

export interface BellaRawResponse {
  mail_template: string;
  messageUuid: string;
  language: string;
  prompt?: string;
  timers?: Record<string, number>;
}

export interface AgentResponse {
  draft: string;
  enquiries: string;
  unanswered: string;
}

export interface ProductCode {
  value: string;
  label: string;
}

export type UIPanel = "auth-loading" | "access-denied" | "main-panel" | "loading" | "result";
