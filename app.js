// ─────────────────────────────────────────────
//  Configuration — update these before connecting the real API
// ─────────────────────────────────────────────
const BELLA_MCP_URL = "https://YOUR_MCP_URL/generate"; // replace with your MCP Connector URL
const BELLA_API_TOKEN = "YOUR_TOKEN";                   // replace with your API token
const SIMULATE_API = true; // set to false when the Bella API is connected

// ─────────────────────────────────────────────
//  Office.js initialization
// ─────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btn-generate").addEventListener("click", generateResponse);
    document.getElementById("btn-insert").addEventListener("click", insertDraft);
    document.getElementById("btn-regenerate").addEventListener("click", generateResponse);
  }
});

// ─────────────────────────────────────────────
//  Read the email and trigger Bella
// ─────────────────────────────────────────────
function generateResponse() {
  hideResult();
  hideError();
  showLoading(true);
  setGenerateButtonEnabled(false);

  var item = Office.context.mailbox.item;
  var emailSubject = item.subject || "(no subject)";
  var emailFrom = item.from ? item.from.emailAddress : "unknown sender";

  item.body.getAsync("text", function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      showLoading(false);
      setGenerateButtonEnabled(true);
      showError("Unable to read the email content. Please try again.");
      return;
    }

    var emailBody = result.value || "";

    callBellaAPI(emailBody, emailSubject, emailFrom)
      .then(function (draft) {
        showLoading(false);
        setGenerateButtonEnabled(true);
        showResult(draft);
      })
      .catch(function (err) {
        showLoading(false);
        setGenerateButtonEnabled(true);
        showError(err.message || "An unexpected error occurred.");
      });
  });
}

// ─────────────────────────────────────────────
//  Call the Bella API (or simulate)
// ─────────────────────────────────────────────
function callBellaAPI(emailBody, emailSubject, emailFrom) {
  if (SIMULATE_API) {
    return new Promise(function (resolve) {
      setTimeout(function () {
        var simulatedDraft =
          "Dear Guest,\n\n" +
          "Thank you for your message regarding \"" + emailSubject + "\".\n\n" +
          "We have received your inquiry and our team is looking into it with the utmost attention. " +
          "We will get back to you as soon as possible with a full response.\n\n" +
          "In the meantime, please do not hesitate to reach out if you have any additional questions.\n\n" +
          "Warm regards,\n\n" +
          "Belmond Guest Relations";
        resolve(simulatedDraft);
      }, 1500);
    });
  }

  // ── Real request to the MCP Connector (uncomment when the API is available) ──
  return fetch(BELLA_MCP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + BELLA_API_TOKEN
    },
    body: JSON.stringify({
      emailBody: emailBody,
      emailSubject: emailSubject,
      emailFrom: emailFrom
    })
  })
    .then(function (response) {
      if (!response.ok) {
        return response.text().then(function (text) {
          throw new Error("API error (" + response.status + "): " + (text || response.statusText));
        });
      }
      return response.json();
    })
    .then(function (data) {
      if (!data || !data.draft) {
        throw new Error("Invalid or empty response from Bella.");
      }
      return data.draft;
    });
}

// ─────────────────────────────────────────────
//  Insert the draft into the reply area
// ─────────────────────────────────────────────
function insertDraft() {
  var draftText = document.getElementById("draft-text").value;
  if (!draftText.trim()) {
    showError("The draft is empty. Please generate a reply first.");
    return;
  }

  Office.context.mailbox.item.body.setAsync(
    draftText,
    { coercionType: Office.CoercionType.Text },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        showError("Unable to insert the draft. Make sure you are in reply mode.");
      }
    }
  );
}

// ─────────────────────────────────────────────
//  UI helpers
// ─────────────────────────────────────────────
function showLoading(visible) {
  var el = document.getElementById("loading");
  el.classList.toggle("visible", visible);
}

function showResult(draft) {
  document.getElementById("draft-text").value = draft;
  document.getElementById("result").classList.add("visible");
}

function hideResult() {
  document.getElementById("result").classList.remove("visible");
  document.getElementById("draft-text").value = "";
}

function showError(message) {
  document.getElementById("error-message").textContent = message;
  document.getElementById("error").classList.add("visible");
}

function hideError() {
  document.getElementById("error").classList.remove("visible");
  document.getElementById("error-message").textContent = "";
}

function setGenerateButtonEnabled(enabled) {
  document.getElementById("btn-generate").disabled = !enabled;
}
