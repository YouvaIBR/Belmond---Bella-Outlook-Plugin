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

// Fields returned by the prefill endpoint: the AI's best match for each form
// field, as ids (selects) + free-text values. Empty string = no confident match.
export interface PrefillData {
  requisition_type_id: string;
  company_id: string;
  currency_id: string;
  requester_id: string;
  ship_to_contact_id: string;
  business_unit_id: string;
  cost_center_id: string;
  spend_category_id: string;
  supplier_id: string;
  unit_of_measure_id: string;
  item_name: string;
  unit_cost: string;
  quantity: string;
  supplier_item_identifier: string;
  memo: string;
  high_priority: boolean;
}

export interface PrefillResponse {
  success: boolean;
  data: PrefillData;
}

export type UIPanel = "auth-loading" | "access-denied" | "main-panel" | "loading" | "result";

export type WorkdayPanel =
  | "auth-loading"
  | "access-denied"
  | "workday-actions"
  | "workday-form"
  | "loading"
  | "result";
