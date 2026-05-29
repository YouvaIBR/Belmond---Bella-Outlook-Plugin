import type { AgentResponse } from "../types/index.js";

const ALLOWED_TAGS = new Set([
  "p", "br", "strong", "em", "b", "i", "ul", "ol", "li",
  "h1", "h2", "h3", "h4", "div", "span", "a", "blockquote",
]);

// Strips disallowed tags and on* / javascript: attributes before any DOM insertion.
export function sanitizeHtml(html: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  removeDisallowedNodes(doc.body);
  return doc.body.innerHTML;
}

function removeDisallowedNodes(node: Element): void {
  for (const child of Array.from(node.childNodes)) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;

    const el = child as Element;
    const tag = el.tagName.toLowerCase();

    if (!ALLOWED_TAGS.has(tag)) {
      node.removeChild(el);
      continue;
    }

    for (const attr of Array.from(el.attributes)) {
      if (attr.name.startsWith("on")) el.removeAttribute(attr.name);
    }

    if (tag === "a") {
      const href = el.getAttribute("href") ?? "";
      if (href.toLowerCase().startsWith("javascript:")) el.removeAttribute("href");
    }

    removeDisallowedNodes(el);
  }
}

function setSafeHtml(id: string, html: string): void {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = sanitizeHtml(html); // nosec: sanitized above
}

export function renderResult(response: AgentResponse): void {
  setSafeHtml("draft-preview", response.draft);
  setSafeHtml("enquiries", response.enquiries);
  setSafeHtml("unanswered", response.unanswered);
}

export function getDraftHtml(): string {
  return document.getElementById("draft-preview")?.innerHTML ?? "";
}
