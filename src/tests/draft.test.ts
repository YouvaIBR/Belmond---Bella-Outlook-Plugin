import { beforeEach, describe, expect, it } from "vitest";
import { getDraftHtml, renderResult, sanitizeHtml } from "../ui/draft.js";

describe("sanitizeHtml", () => {
  it("passes safe tags through", () => {
    expect(sanitizeHtml("<p>Hello <strong>world</strong></p>")).toBe(
      "<p>Hello <strong>world</strong></p>",
    );
  });

  it("removes script tags", () => {
    const result = sanitizeHtml("<p>Text</p><script>alert(\'xss\')</script>");
    expect(result).not.toContain("<script>");
    expect(result).toContain("<p>Text</p>");
  });

  it("removes on* event attributes", () => {
    const result = sanitizeHtml("<p onclick=\"alert(1)\">Click</p>");
    expect(result).not.toContain("onclick");
    expect(result).toContain("<p>");
  });

  it("removes disallowed tags (img, iframe)", () => {
    const result = sanitizeHtml("<img src=\"x\" onerror=\"alert(1)\"/><iframe src=\"x\"></iframe>");
    expect(result).not.toContain("img");
    expect(result).not.toContain("iframe");
  });

  it("removes javascript: href from anchors", () => {
    const result = sanitizeHtml("<a href=\"javascript:alert(1)\">link</a>");
    expect(result).not.toContain("javascript:");
    expect(result).toContain("<a");
  });

  it("keeps valid href on anchors", () => {
    const result = sanitizeHtml("<a href=\"https://example.com\">link</a>");
    expect(result).toContain("https://example.com");
  });

  it("returns empty string for empty input", () => {
    expect(sanitizeHtml("")).toBe("");
  });
});

describe("renderResult", () => {
  beforeEach(() => {
    document.body.innerHTML =
      "<div id=\"draft-preview\"></div><div id=\"enquiries\"></div><div id=\"unanswered\"></div>";
  });

  it("populates all three DOM sections", () => {
    renderResult({ draft: "<p>Draft</p>", enquiries: "<p>Q1</p>", unanswered: "<p>Q2</p>" });
    expect(document.getElementById("draft-preview")?.textContent).toContain("Draft");
    expect(document.getElementById("enquiries")?.textContent).toContain("Q1");
    expect(document.getElementById("unanswered")?.textContent).toContain("Q2");
  });

  it("does not throw when an element is missing", () => {
    document.body.innerHTML = "<div id=\"draft-preview\"></div>";
    expect(() =>
      renderResult({ draft: "<p>ok</p>", enquiries: "<p>ok</p>", unanswered: "<p>ok</p>" }),
    ).not.toThrow();
  });
});

describe("getDraftHtml", () => {
  it("returns innerHTML of draft-preview", () => {
    document.body.innerHTML = "<div id=\"draft-preview\"><p>My draft</p></div>";
    expect(getDraftHtml()).toBe("<p>My draft</p>");
  });

  it("returns empty string when element is absent", () => {
    document.body.innerHTML = "";
    expect(getDraftHtml()).toBe("");
  });
});
