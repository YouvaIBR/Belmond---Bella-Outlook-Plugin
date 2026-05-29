import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AgentError, callAgent, parseBellaResponse } from "../api/agent.js";
import type { AgentRequest, BellaRawResponse } from "../types/index.js";

const BASE_URL = "https://n8n.test.internal/webhook/bella";

vi.mock("../auth/msal.js", () => ({
  acquireToken: vi.fn().mockResolvedValue("mock-token"),
}));

const mockRequest: AgentRequest = {
  emailBody: "I would like to book a room",
  emailSubject: "Booking enquiry",
  emailFrom: "guest@example.com",
  productCode: "LRS",
};

describe("callAgent", () => {
  beforeEach(() => {
    vi.stubGlobal("fetch", vi.fn());
  });

  afterEach(() => {
    vi.clearAllMocks();
  });

  it("returns parsed response on 200", async () => {
    const rawResponse: BellaRawResponse = {
      mail_template: "<b>Enquiries</b>:<br>Book a room<br><b>Unanswered</b>:<br>Check-in date<br><b>Email</b>:<br><hr><p>Thank you</p>",
      messageUuid: "abc-123",
      language: "en_US",
    };
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: { get: () => "application/json" },
      text: () => Promise.resolve(JSON.stringify(rawResponse)),
    } as unknown as Response);

    const result = await callAgent(mockRequest, BASE_URL);
    expect(result.draft).toContain("Thank you");
    expect(result.enquiries).toContain("Book a room");
    expect(result.unanswered).toContain("Check-in date");
    expect(fetch).toHaveBeenCalledWith(
      BASE_URL,
      expect.objectContaining({
        method: "POST",
        headers: expect.objectContaining({ Authorization: "Bearer mock-token" }),
      }),
    );
  });

  it("throws AgentError on 401", async () => {
    vi.mocked(fetch).mockResolvedValue({
      ok: false,
      status: 401,
      statusText: "Unauthorized",
      headers: { get: () => null },
      text: () => Promise.resolve(""),
    } as unknown as Response);

    await expect(callAgent(mockRequest, BASE_URL)).rejects.toThrow(AgentError);
    await expect(callAgent(mockRequest, BASE_URL)).rejects.toMatchObject({ statusCode: 401 });
  });

  it("throws AgentError with body message on 502", async () => {
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: false,
      status: 502,
      statusText: "Bad Gateway",
      headers: { get: () => null },
      text: () => Promise.resolve("n8n unavailable"),
    } as unknown as Response);

    await expect(callAgent(mockRequest, BASE_URL)).rejects.toMatchObject({
      statusCode: 502,
      message: "n8n unavailable",
    });
  });

  it("propagates network errors", async () => {
    vi.mocked(fetch).mockRejectedValueOnce(new TypeError("Failed to fetch"));
    await expect(callAgent(mockRequest, BASE_URL)).rejects.toThrow("Failed to fetch");
  });

  it("sends the Bearer token in the Authorization header", async () => {
    vi.mocked(fetch).mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: { get: () => "application/json" },
      text: () => Promise.resolve(JSON.stringify({
        mail_template: "<b>Email</b>:<br><hr><p>ok</p>",
        messageUuid: "x",
        language: "en_US",
      })),
    } as unknown as Response);

    await callAgent(mockRequest, BASE_URL);

    const [, options] = vi.mocked(fetch).mock.calls[0] ?? [];
    const headers = options?.headers as Record<string, string>;
    expect(headers?.Authorization).toBe("Bearer mock-token");
  });
});

describe("parseBellaResponse", () => {
  const base: BellaRawResponse = {
    mail_template: "<b>Enquiries</b>:<br>Book a room<br><b>Unanswered</b>:<br>Check-in date<br><b>Email</b>:<br><hr><p>Thank you for your enquiry.</p>",
    messageUuid: "abc-123",
    language: "en_US",
  };

  it("extracts draft from Email section", () => {
    const result = parseBellaResponse(base);
    expect(result.draft).toBe("<p>Thank you for your enquiry.</p>");
  });

  it("extracts enquiries section", () => {
    const result = parseBellaResponse(base);
    expect(result.enquiries).toContain("Book a room");
  });

  it("extracts unanswered section", () => {
    const result = parseBellaResponse(base);
    expect(result.unanswered).toContain("Check-in date");
  });

  it("falls back to full template when Email section is missing", () => {
    const result = parseBellaResponse({ ...base, mail_template: "<p>fallback</p>" });
    expect(result.draft).toBe("<p>fallback</p>");
  });

  it("returns empty strings when sections are missing", () => {
    const result = parseBellaResponse({ ...base, mail_template: "" });
    expect(result.enquiries).toBe("");
    expect(result.unanswered).toBe("");
    expect(result.draft).toBe("");
  });
});
