import { sanitizeHtml } from "../ui/draft.js";

Office.onReady(() => {
  const draft = localStorage.getItem("bella_draft") ?? "";
  const editor = document.getElementById("editor") as HTMLDivElement;
  editor.innerHTML = sanitizeHtml(draft);

  document.getElementById("copy-btn")?.addEventListener("click", handleCopy);
  document.getElementById("close-btn")?.addEventListener("click", () => {
    Office.context.ui.messageParent("close");
  });
});

async function handleCopy(): Promise<void> {
  const editor = document.getElementById("editor") as HTMLDivElement;
  const text = editor.innerText;

  try {
    await navigator.clipboard.writeText(text);
    showToast("Copied to clipboard");
  } catch {
    // Fallback for browsers that block clipboard without user gesture
    const range = document.createRange();
    range.selectNodeContents(editor);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    document.execCommand("copy");
    sel?.removeAllRanges();
    showToast("Copied to clipboard");
  }
}

function showToast(message: string): void {
  const toast = document.getElementById("toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2000);
}
