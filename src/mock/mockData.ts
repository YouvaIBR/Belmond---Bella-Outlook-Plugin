import type { AgentResponse, MailItem, WorkdayOptions } from "../types/index.js";

export const mockUserName = "Youva Ibrahim";

export const mockMailItem: MailItem = {
  subject: "Restaurant reservation",
  from: "sarah.thompson@gmail.com",
  body: "Dear La Residencia Team,I hope this message finds you well. I am writing to inquire about availability at your property for an upcoming trip to Mallorca.We are a group of 2 adults looking to stay at La Residencia from July 14 to July 21, 2026 (7 nights). Could you please let us know if you have any rooms or suites available during that period, along with the current rates?Additionally, I would love to know whether it is possible to arrange a private boat rental through the hotel, either for a half-day or full-day excursion along the coast. If so, could you provide details on the options available, pricing, and whether a captain/crew is included?We would very much appreciate any information you can share, and we look forward to hearing from you.Warm regards,Sarah Thompson +1 (917) 555-0284 sarah.thompson@gmail.com",
};

export const mockWorkdayOptions: WorkdayOptions = {
  requisitionType: [
    { id: "edac21e76a421000907d48e59cb10000", descriptor: "3 Design Studio" },
    { id: "edac21e76a421000907d48e59cb10001", descriptor: "Operating Supplies" },
  ],
  company: [
    { id: "b5cdd0d7b13b0130ac4a135226183e7f", descriptor: "Son Moragues SA" },
    { id: "b5cdd0d7b13b0130ac4a135226183e80", descriptor: "Belmond Management Ltd" },
  ],
  currency: [
    { id: "eae312fc5152410cb4c8b452c26320a6", descriptor: "Euro" },
    { id: "eae312fc5152410cb4c8b452c26320a7", descriptor: "Pound Sterling" },
  ],
  requester: [
    { id: "5c756a2e52f3015556d8ba403e2e4300", descriptor: "Mauro Giannini" },
    { id: "5c756a2e52f3015556d8ba403e2e4301", descriptor: "Lizelle Bezuidenhout" },
  ],
  shipToContact: [
    { id: "b233c2d5136501819817e9a7a72aef3a", descriptor: "Lizelle Bezuidenhout" },
    { id: "b233c2d5136501819817e9a7a72aef3b", descriptor: "Mauro Giannini" },
  ],
  businessUnit: [
    { id: "b5cdd0d7b13b016891178d1d9117cf36", descriptor: "Business Unit: La Residencia" },
    { id: "b5cdd0d7b13b016891178d1d9117cf37", descriptor: "Business Unit: Le Manoir" },
  ],
  costCenter: [
    { id: "b5cdd0d7b13b01a54f1579c82918bb87", descriptor: "LRSNON Non Departmental" },
    { id: "b5cdd0d7b13b01a54f1579c82918bb88", descriptor: "LRSFB Food & Beverage" },
  ],
  spendCategory: [
    { id: "dda1cdac393b01b8b11b1eb83e0a4715", descriptor: "Contract Services Architecture" },
    { id: "dda1cdac393b01b8b11b1eb83e0a4716", descriptor: "Operating Supplies" },
  ],
  supplier: [
    { id: "a50b87eaeb9d100190b29a633b7d0000", descriptor: "Fabio Maggioni" },
    { id: "a50b87eaeb9d100190b29a633b7d0001", descriptor: "Mallorca Construction SL" },
  ],
  unitOfMeasure: [
    { id: "b5cdd0d7b13b0176d3827b4f2018666a", descriptor: "Each" },
    { id: "b5cdd0d7b13b0176d3827b4f2018666b", descriptor: "Hour" },
  ],
};

export const mockRequisitionResponse =
  "<p>Your requisition <b>REQ-2026-00481</b> has been created and submitted for approval.</p><p>It has been routed to <b>Mauro Giannini</b> for review. The attached document(s) were linked to the requisition line.</p><p>You'll receive a notification once it has been approved.</p>";

export const mockAgentResponse: AgentResponse = {
  draft: "<p>Dear Sarah,</p><p>Thank you for your email and your interest in La Residencia, a Belmond Hotel.</p><p>Regarding your inquiry about availability from <b>July 14 to July 21, 2026</b> for <b>2 adults</b>, <span style=\"background-color: green\">we are currently checking our system for room and suite availability and will get back to you shortly with the rates</span>.</p><p>We look forward to helping you plan a memorable stay with us.</p><p>Warm regards,</p><br>Concierge Team<br>La Residencia, A Belmond Hotel",
  enquiries: "<ul><li>Check availability for 2 adults from July 14 to July 21, 2026 (7 nights) and provide rates.</li><li>Provide details on private boat rental options, pricing, and whether a captain/crew is included.</li></ul>",
  unanswered: "",
};
