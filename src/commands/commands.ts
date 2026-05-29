import { AgentError, callAgent } from "../api/agent.js";
import { acquireToken } from "../auth/msal.js";
import { getMailItem } from "../services/office.js";

const NOTIFICATION_ID = "bella-status";
const SESSION_KEY = "bella_draft";
const SESSION_ERROR_KEY = "bella_draft_error";

function bellaReply(event: Office.AddinCommands.Event): void {
  const item = Office.context.mailbox.item as Office.MessageRead | null;
  if (!item) {
    event.completed();
    return;
  }

  // Clear previous draft
  localStorage.removeItem(SESSION_KEY);
  localStorage.removeItem(SESSION_ERROR_KEY);

  // Open the reply form immediately (must be synchronous, within user gesture)
  // The OnNewMessageCompose handler will replace this placeholder with the real draft
  item.displayReplyAllForm({
    htmlBody: `<p style="color:#888;font-style:italic;">Bella is thinking…</p>`,
  });

  // API call runs in background — event.completed() deferred until done
  void runAgentAndStore(event, item);
}

async function runAgentAndStore(
  event: Office.AddinCommands.Event,
  item: Office.MessageRead,
): Promise<void> {
  try {
    await acquireToken();
    const mailItem = await getMailItem();
    const response = await callAgent({
      emailBody: mailItem.body,
      emailSubject: mailItem.subject,
      emailFrom: mailItem.from,
      productCode: "LRS",
    });

    // Store draft so the OnNewMessageCompose handler can read it
    localStorage.setItem(SESSION_KEY, response.draft);
  } catch (err) {
    const message =
      err instanceof AgentError
        ? `Bella error ${err.statusCode.toString()}: ${err.message}. Please retry.`
        : err instanceof Error
          ? `${err.message} — Please retry.`
          : "Bella encountered an error. Please retry.";

    localStorage.setItem(SESSION_ERROR_KEY, message);
    await notify(item, "errorMessage", message, true);
  } finally {
    event.completed();
  }
}

function notify(
  item: Office.MessageRead,
  type: "informationalMessage" | "errorMessage",
  message: string,
  persistent: boolean,
): Promise<void> {
  return new Promise((resolve) => {
    item.notificationMessages.replaceAsync(
      NOTIFICATION_ID,
      {
        type:
          type === "errorMessage"
            ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
            : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message,
        icon: "Icon.16x16",
        persistent,
      },
      () => resolve(),
    );
  });
}

Office.onReady(() => {
  Office.actions.associate("bellaReply", bellaReply);
});
