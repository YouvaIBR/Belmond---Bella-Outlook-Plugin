// Reads real (non-inline) PDF attachments from the current mail item and returns
// them as File objects. Office.js hands us base64 content, which we decode to a
// Blob immediately so no base64 ever leaves the browser — the network payload is
// pure binary via multipart/form-data.

const PDF_CONTENT_TYPE = "application/pdf";

function getAttachments(): Promise<Office.AttachmentDetails[]> {
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

    // attachments is available synchronously in read mode.
    const attachments = item.attachments ?? [];
    resolve(attachments);
  });
}

function isPdfAttachment(att: Office.AttachmentDetails): boolean {
  // Keep only real file attachments (skip inline images / item attachments).
  if (att.attachmentType !== Office.MailboxEnums.AttachmentType.File) return false;
  if (att.isInline) return false;

  const contentTypeIsPdf = att.contentType?.toLowerCase() === PDF_CONTENT_TYPE;
  const nameIsPdf = att.name?.toLowerCase().endsWith(".pdf") ?? false;
  return contentTypeIsPdf || nameIsPdf;
}

function getAttachmentFile(att: Office.AttachmentDetails): Promise<File> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item?.getAttachmentContentAsync) {
      reject(new Error("Attachment content is not available for this item."));
      return;
    }

    item.getAttachmentContentAsync(att.id, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      const content = result.value;
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        reject(new Error(`Unexpected attachment format: ${content.format}`));
        return;
      }

      try {
        const blob = base64ToBlob(content.content, PDF_CONTENT_TYPE);
        resolve(new File([blob], att.name, { type: PDF_CONTENT_TYPE }));
      } catch (err) {
        reject(err instanceof Error ? err : new Error("Failed to decode attachment."));
      }
    });
  });
}

function base64ToBlob(base64: string, contentType: string): Blob {
  const byteChars = atob(base64);
  const byteNumbers = new Uint8Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) {
    byteNumbers[i] = byteChars.charCodeAt(i);
  }
  return new Blob([byteNumbers], { type: contentType });
}

// Returns all real PDF attachments of the current mail item as binary File objects.
export async function getPdfAttachments(): Promise<File[]> {
  const attachments = await getAttachments();
  const pdfs = attachments.filter(isPdfAttachment);
  return Promise.all(pdfs.map(getAttachmentFile));
}
