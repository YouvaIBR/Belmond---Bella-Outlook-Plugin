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

// Workday entities are always an { id, descriptor } pair. The descriptor is the
// human-readable label shown in the UI; the id is what Workday persists.
export interface WorkdayOption {
  id: string;
  descriptor: string;
}

// All the option lists the form needs, returned by the Workday data GET endpoint.
// Every list feeds one <select>.
export interface WorkdayOptions {
  requisitionType: WorkdayOption[];
  company: WorkdayOption[];
  currency: WorkdayOption[];
  requester: WorkdayOption[];
  shipToContact: WorkdayOption[];
  businessUnit: WorkdayOption[];
  costCenter: WorkdayOption[];
  spendCategory: WorkdayOption[];
  supplier: WorkdayOption[];
  unitOfMeasure: WorkdayOption[];
}

export interface WorkdayDataResponse {
  success: boolean;
  data: WorkdayOptions;
}

// The requisition form is submitted as a single free-text `query` (all selected
// ids + labels + free-text values + the email content) plus the binary
// documents. The backend agent parses the query, so no structured payload type
// is needed on the client.

export type UIPanel = "auth-loading" | "access-denied" | "main-panel" | "loading" | "result";

export type WorkdayPanel =
  | "auth-loading"
  | "access-denied"
  | "workday-actions"
  | "workday-form"
  | "loading"
  | "result";
