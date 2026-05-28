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

    item.displayReplyFormAsync({ htmlBody: response.draft });
  } catch (err) {
    const message =
      err instanceof AgentError
        ? `Bella error ${err.statusCode.toString()}: ${err.message}. Please retry.`
        : err instanceof Error
          ? `${err.message} — Please retry.`
          : "Bella encountered an error. Please retry.";

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
