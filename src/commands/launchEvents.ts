const SESSION_KEY = "bella_draft";
const SESSION_ERROR_KEY = "bella_draft_error";

// Fired when the reply form opened by displayReplyAllForm becomes ready.
// Reads the draft stored in localStorage by bellaReply and injects it into the body.
function onNewMessageCompose(event: Office.MailboxEvent): void {
  const item = Office.context.mailbox.item as Office.MessageCompose | null;
  if (!item) {
    event.completed();
    return;
  }

  void injectDraft(event, item);
}

async function injectDraft(
  event: Office.MailboxEvent,
  item: Office.MessageCompose,
): Promise<void> {
  const draft = localStorage.getItem(SESSION_KEY);

  if (draft) {
    await setBodyAsync(item, draft);
    localStorage.removeItem(SESSION_KEY);
  } else {
    // Draft not ready yet — poll while API is still in flight
    const result = await pollForDraft();
    if (result) {
      await setBodyAsync(item, result);
    }
    // If still nothing, "Bella is thinking…" placeholder stays;
    // error notification from commands.ts informs the user.
  }

  event.completed();
}

function setBodyAsync(item: Office.MessageCompose, html: string): Promise<void> {
  return new Promise((resolve) => {
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, () => resolve());
  });
}

// Poll up to 30 s in 1 s increments waiting for the API call to finish.
async function pollForDraft(): Promise<string | null> {
  for (let i = 0; i < 30; i++) {
    await sleep(1000);
    const draft = localStorage.getItem(SESSION_KEY);
    if (draft) {
      localStorage.removeItem(SESSION_KEY);
      return draft;
    }
    if (localStorage.getItem(SESSION_ERROR_KEY)) return null;
  }
  return null;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

Office.onReady(() => {
  Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
});
