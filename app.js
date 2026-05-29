// ─────────────────────────────────────────────
//  Configuration
// ─────────────────────────────────────────────
const BELLA_MCP_URL =
  "https://n8n-985953980051.europe-west1.run.app/webhook/bella/concierge";
const BELLA_API_TOKEN = "TOKEN";
const BELLA_API_COOKIE = "GAESA=THE_COOKIE";

// Azure AD — replace with real values before production deployment
const AZURE_CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID";
const AZURE_TENANT_ID = "YOUR_TENANT_ID";

// Object IDs of the Azure AD groups authorized to use this add-in
// Found in: Entra ID → Groups → select group → Overview → Object ID
const AUTHORIZED_GROUP_IDS = ["YOUR_AUTHORIZED_GROUP_OBJECT_ID"];

// Set to false to enable the Azure AD group membership check
const SKIP_ACCESS_CONTROL = true;

// ─────────────────────────────────────────────
//  Office.js initialization
// ─────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    runAccessControl();
  }
});

// ─────────────────────────────────────────────
//  Access control gate
// ─────────────────────────────────────────────
async function runAccessControl() {
  if (SKIP_ACCESS_CONTROL) {
    showMainPanel();
    return;
  }

  showAuthLoading(true);

  try {
    const token = await getGraphToken();
    const isAuthorized = await checkGroupMembership(token);

    showAuthLoading(false);

    if (isAuthorized) {
      showMainPanel();
    } else {
      showAccessDenied(
        "You are not authorized to use this add-in. Please contact your administrator."
      );
    }
  } catch (err) {
    showAuthLoading(false);
    showAccessDenied(
      "Authentication failed: " + (err.message || "unknown error")
    );
  }
}

// ─────────────────────────────────────────────
//  MSAL — acquire a Microsoft Graph token via NAA
//  Requires @azure/msal-browser v3+ loaded in index.html
// ─────────────────────────────────────────────
async function getGraphToken() {
  const msalConfig = {
    auth: {
      clientId: AZURE_CLIENT_ID,
      authority: "https://login.microsoftonline.com/" + AZURE_TENANT_ID,
      supportsNestedAppAuth: true,
    },
    cache: {
      cacheLocation: "localStorage",
    },
  };

  // createNestablePublicClientApplication is the NAA entry point (MSAL v3+)
  const msalInstance = await msal.createNestablePublicClientApplication(
    msalConfig
  );

  const tokenRequest = {
    scopes: [
      "https://graph.microsoft.com/User.Read",
      "https://graph.microsoft.com/GroupMember.Read.All",
    ],
  };

  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      tokenRequest.account = accounts[0];
    }
    const response = await msalInstance.acquireTokenSilent(tokenRequest);
    return response.accessToken;
  } catch (err) {
    // Silent acquisition failed — prompt the user interactively
    const response = await msalInstance.acquireTokenPopup(tokenRequest);
    return response.accessToken;
  }
}

// ─────────────────────────────────────────────
//  Microsoft Graph — check if user belongs to an authorized group
//  POST /me/checkMemberGroups
//  Requires delegated permission: GroupMember.Read.All (admin consent needed)
// ─────────────────────────────────────────────
async function checkGroupMembership(accessToken) {
  const response = await fetch(
    "https://graph.microsoft.com/v1.0/me/checkMemberGroups",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ groupIds: AUTHORIZED_GROUP_IDS }),
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error("Graph API error (" + response.status + "): " + text);
  }

  const data = await response.json();
  // data.value contains the subset of AUTHORIZED_GROUP_IDS the user belongs to
  return data.value.length > 0;
}

// ─────────────────────────────────────────────
//  Main add-in logic (only reached after access control passes)
// ─────────────────────────────────────────────
function showMainPanel() {
  document.getElementById("auth-loading").style.display = "none";
  document.getElementById("access-denied").style.display = "none";
  document.getElementById("main-panel").style.display = "block";

  document
    .getElementById("btn-generate")
    .addEventListener("click", generateResponse);
  document.getElementById("btn-insert").addEventListener("click", insertDraft);
  document
    .getElementById("btn-regenerate")
    .addEventListener("click", generateResponse);
}

async function generateResponse() {
  hideResult();
  hideError();
  showLoading(true);
  setGenerateButtonEnabled(false);

  var item = Office.context.mailbox.item;
  var emailSubject = item.subject || "(no subject)";
  var emailFrom = item.from ? item.from.emailAddress : "unknown sender";

  item.body.getAsync("text", async function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      showLoading(false);
      setGenerateButtonEnabled(true);
      showError("Unable to read the email content. Please try again.");
      return;
    }

    var emailBody = result.value || "";

    await callBellaAPI(emailBody, emailSubject, emailFrom)
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

async function callBellaAPI(emailBody, emailSubject, emailFrom) {
  return await fetch(BELLA_MCP_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Basic " + BELLA_API_TOKEN,
      Cookie: BELLA_API_COOKIE,
    },
    body: JSON.stringify({
      query: emailBody,
      product_code: "LRS",
      vision: false,
    }),
  })
    .then(async function (response) {
      if (!response.ok) {
        return response.text().then(function (text) {
          throw new Error(
            "API error (" +
              response.status +
              "): " +
              (text || response.statusText)
          );
        });
      }
      return await response.json();
    })
    .then(function (data) {
      if (!data || data.length === 0) {
        throw new Error("Invalid or empty response from Bella.");
      }
      return data[0].mail_template;
    });
}

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
        showError(
          "Unable to insert the draft. Make sure you are in reply mode."
        );
      }
    }
  );
}

// ─────────────────────────────────────────────
//  UI helpers
// ─────────────────────────────────────────────
function showAuthLoading(visible) {
  document.getElementById("auth-loading").style.display = visible
    ? "flex"
    : "none";
}

function showAccessDenied(message) {
  var el = document.getElementById("access-denied");
  el.querySelector(".access-denied-message").textContent = message;
  el.style.display = "block";
}

function showLoading(visible) {
  document.getElementById("loading").classList.toggle("visible", visible);
}

// Strip script/event-handler attributes from an HTML string before rendering.
function sanitizeHtml(html) {
  var tmp = document.createElement("div");
  // Use the browser's HTML parser — no script tags, remove on* attributes
  tmp.innerHTML = html.replace(/<script[\s\S]*?<\/script>/gi, "");
  tmp.querySelectorAll("*").forEach(function (el) {
    Array.from(el.attributes).forEach(function (attr) {
      if (
        /^on/i.test(attr.name) ||
        (attr.name.toLowerCase() === "href" && /^javascript:/i.test(attr.value))
      ) {
        el.removeAttribute(attr.name);
      }
    });
  });
  return tmp.innerHTML;
}

// Parse the HTML response from the API into three sections.
// Expected format: <b>Enquiries</b>:<br>...<b>Unanswered</b>:<br>...<b>Email</b>:<br><hr>...
function parseApiResponse(html) {
  var enquiriesMatch = html.match(
    /<b>Enquiries<\/b>:([\s\S]*?)(?=<b>Unanswered<\/b>|$)/i
  );
  var unansweredMatch = html.match(
    /<b>Unanswered<\/b>:([\s\S]*?)(?=<b>Email<\/b>|$)/i
  );
  var emailMatch = html.match(/<b>Email<\/b>:<br\s*\/?><hr\s*\/?>([\s\S]*?)$/i);

  return {
    enquiries: enquiriesMatch
      ? enquiriesMatch[1].replace(/^<br\s*\/?>/i, "").trim()
      : "",
    unanswered: unansweredMatch
      ? unansweredMatch[1].replace(/^<br\s*\/?>/i, "").trim()
      : "",
    email: emailMatch ? emailMatch[1].trim() : "",
  };
}

function showResult(rawHtml) {
  var parsed = parseApiResponse(rawHtml);
  // Draft preview
  var emailHtml = sanitizeHtml(parsed.email || rawHtml);
  document.getElementById("draft-preview").innerHTML = emailHtml;

  // Hidden textarea: plain text for insertion into Outlook
  var tmp = document.createElement("div");
  tmp.innerHTML = emailHtml;
  document.getElementById("draft-text").value =
    tmp.innerText || tmp.textContent || "";

  document.getElementById("result").classList.add("visible");
}

function hideResult() {
  document.getElementById("result").classList.remove("visible");
  document.getElementById("draft-preview").innerHTML = "";
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
