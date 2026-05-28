import { AgentError, callAgent } from "../api/agent.js";
import { acquireToken } from "../auth/msal.js";
import { getMailItem } from "../services/office.js";

const NOTIFICATION_ID = "bella-status";

async function bellaReply(event: Office.AddinCommands.Event): Promise<void> {
  const item = Office.context.mailbox.item as Office.MessageRead | null;
  if (!item) {
    event.completed();
    return;
  }

  await notify(item, "informationalMessage", "Bella is thinking…", false);

  try {
    // Ensure token is ready before the heavy API call
    await acquireToken();

    const mailItem = await getMailItem();
    const response = await callAgent({
      emailBody: mailItem.body,
      emailSubject: mailItem.subject,
      emailFrom: mailItem.from,
      productCode: "LRS",
    });

    item.notificationMessages.removeAsync(NOTIFICATION_ID);

    localStorage.setItem("bella_draft", response.draft);

    Office.context.ui.displayDialogAsync(
      `${window.location.origin}/Belmond---Bella-Outlook-Plugin/reply-dialog.html`,
      { height: 60, width: 40, promptBeforeOpen: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          void notify(item, "errorMessage", "Could not open the draft dialog. Please retry.", true);
          event.completed();
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
          dialog.close();
          event.completed();
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          // User closed the dialog manually
          event.completed();
        });
      },
    );
  } catch (err) {
    const message =
      err instanceof AgentError
        ? `Bella error ${err.statusCode.toString()}: ${err.message}. Please retry.`
        : err instanceof Error
          ? `${err.message} — Please retry.`
          : "Bella encountered an error. Please retry.";

    await notify(item, "errorMessage", message, true);
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
