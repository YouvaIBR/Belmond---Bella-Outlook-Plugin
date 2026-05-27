import type { AgentResponse, MailItem } from "../types/index.js";

export const mockMailItem: MailItem = {
  subject: "Restaurant reservation",
  from: "sarah.thompson@gmail.com",
  body: "Dear La Residencia Team,I hope this message finds you well. I am writing to inquire about availability at your property for an upcoming trip to Mallorca.We are a group of 2 adults looking to stay at La Residencia from July 14 to July 21, 2026 (7 nights). Could you please let us know if you have any rooms or suites available during that period, along with the current rates?Additionally, I would love to know whether it is possible to arrange a private boat rental through the hotel, either for a half-day or full-day excursion along the coast. If so, could you provide details on the options available, pricing, and whether a captain/crew is included?We would very much appreciate any information you can share, and we look forward to hearing from you.Warm regards,Sarah Thompson +1 (917) 555-0284 sarah.thompson@gmail.com",
};

export const mockAgentResponse: AgentResponse = {
  draft: "<p>Dear Sarah,</p><p>Thank you for your email and your interest in La Residencia, a Belmond Hotel.</p><p>Regarding your inquiry about availability from <b>July 14 to July 21, 2026</b> for <b>2 adults</b>, <span style=\"background-color: green\">we are currently checking our system for room and suite availability and will get back to you shortly with the rates</span>.</p><p>We look forward to helping you plan a memorable stay with us.</p><p>Warm regards,</p><br>Concierge Team<br>La Residencia, A Belmond Hotel",
  enquiries: "<ul><li>Check availability for 2 adults from July 14 to July 21, 2026 (7 nights) and provide rates.</li><li>Provide details on private boat rental options, pricing, and whether a captain/crew is included.</li></ul>",
  unanswered: "",
};
