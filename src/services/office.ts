import type { MailItem } from "../types/index.js";

export function getMailItem(): Promise<MailItem> {
  return new Promise((resolve, reject) => {
    if (typeof Office === "undefined" || !Office.context?.mailbox) {
      reject(new Error("Bella must be opened inside Outlook. Use npm run dev:mock to test locally."));
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No mail item selected. Open an email before using Bella."));
      return;
    }

    const subject = item.subject ?? "";
    const from = item.from?.emailAddress ?? "";

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }
      resolve({ subject, from, body: result.value });
    });
  });
}

export function insertDraftReply(html: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (typeof Office === "undefined" || !Office.context?.mailbox) {
      reject(new Error("Bella must be opened inside Outlook."));
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No mail item selected. Open an email before using Bella."));
      return;
    }

    // In read mode, body.setAsync doesn't exist — open a reply form with the draft pre-filled.
    // In compose mode (e.g. inline reply), body.setAsync is available.
    if (typeof (item.body as Office.Body).setAsync === "function") {
      (item.body as Office.Body).setAsync(html, { coercionType: Office.CoercionType.Html }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }
        resolve();
      });
    } else {
      item.displayReplyForm({ htmlBody: html });
      resolve();
    }
  });
}

// Opens a pre-filled reply draft directly in Outlook (inject mode).
// Uses displayReplyFormAsync when available, falls back to the legacy sync form.
export function openReplyForm(html: string): void {
  if (typeof Office === "undefined" || !Office.context?.mailbox) return;

  const item = Office.context.mailbox.item;
  if (!item) return;

  if (typeof item.displayReplyFormAsync === "function") {
    item.displayReplyFormAsync({ htmlBody: html });
  } else {
    item.displayReplyForm({ htmlBody: html });
  }
}
