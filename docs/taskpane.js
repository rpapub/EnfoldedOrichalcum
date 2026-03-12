/* global Office, msal */

// ── Configuration ─────────────────────────────────────────────────────────────
const CLIENT_ID      = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const AUTHORITY      = "https://login.microsoftonline.com/common";
const GRAPH_SCOPES   = ["Mail.ReadWrite"];
const EXTENSION_NAME = "com.rpapub.emailtriage";
// ─────────────────────────────────────────────────────────────────────────────

const COUNTRIES = [
  ["AUS","Australia"],["AUT","Austria"],["BEL","Belgium"],["BRA","Brazil"],
  ["CAN","Canada"],["CHE","Switzerland"],["CHN","China"],["DEU","Germany"],
  ["DNK","Denmark"],["EGY","Egypt"],["ESP","Spain"],["FIN","Finland"],
  ["FRA","France"],["GBR","United Kingdom"],["IDN","Indonesia"],["IND","India"],
  ["ITA","Italy"],["JPN","Japan"],["KOR","Korea, Republic of"],["MEX","Mexico"],
  ["NGA","Nigeria"],["NLD","Netherlands"],["NOR","Norway"],["POL","Poland"],
  ["PRT","Portugal"],["RUS","Russia"],["SAU","Saudi Arabia"],["SWE","Sweden"],
  ["TUR","Turkey"],["USA","United States"],["ZAF","South Africa"]
];

function populateCountries(selectId) {
  const sel = document.getElementById(selectId);
  sel.innerHTML = '<option value="">— select —</option>';
  COUNTRIES.forEach(([code, name]) => {
    const opt = document.createElement("option");
    opt.value = code;
    opt.textContent = `${code} — ${name}`;
    sel.appendChild(opt);
  });
}

function populateYears(selectId) {
  const sel = document.getElementById(selectId);
  const now = new Date().getFullYear();
  sel.innerHTML = "";
  for (let y = now; y >= now - 5; y--) {
    const opt = document.createElement("option");
    opt.value = y;
    opt.textContent = y;
    sel.appendChild(opt);
  }
  sel.value = now;
}

function autoSelectQuarter(selectId) {
  const month = new Date().getMonth();
  const q = Math.floor(month / 3) + 1;
  document.getElementById(selectId).value = `Q${q}`;
}

function buildOutput() {
  return {
    from:      document.getElementById("from").value,
    subject:   document.getElementById("subject").value,
    reference: document.getElementById("f-reference").value,
    category:  document.getElementById("f-category").value,
    priority:  document.getElementById("f-priority").value,
    owner:     document.getElementById("f-owner").value,
    notes:     document.getElementById("f-notes").value,
    country:   document.getElementById("f-country").value,
    year:      document.getElementById("f-year").value,
    quarter:   document.getElementById("f-quarter").value
  };
}

function initForm() {
  populateCountries("f-country");
  populateYears("f-year");
  autoSelectQuarter("f-quarter");
}

function showStatus(msg, isError) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status " + (isError ? "error" : "ok");
  if (!isError) setTimeout(() => { el.textContent = ""; el.className = "status"; }, 4000);
}

// ── MSAL (silent only — popup is blocked inside Outlook task panes) ───────────
let msalInstance = null;

function initMsal() {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: AUTHORITY,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "sessionStorage" }
  });
}

async function getTokenSilent() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: accounts[0]
    });
    return result.accessToken;
  } catch (_) {
    return null;
  }
}

// ── Office Dialog API auth ────────────────────────────────────────────────────
// Outlook task panes block window.open (MSAL popup), so we use Office Dialog.
// The auth-dialog.html page handles the MSAL redirect flow and posts the token back.
function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    const dialogUrl =
      window.location.origin +
      window.location.pathname.replace("taskpane.html", "auth-dialog.html");

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 30, promptBeforeOpen: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Dialog failed: ${result.error.message}`));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
          dialog.close();
          try {
            const payload = JSON.parse(msg.message);
            if (payload.error) reject(new Error(payload.error));
            else resolve(payload.accessToken);
          } catch (e) {
            reject(new Error("Invalid message from auth dialog"));
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
          if (evt.error === 12006) reject(new Error("Auth dialog closed by user."));
        });
      }
    );
  });
}

async function getToken() {
  const silent = await getTokenSilent();
  if (silent) return silent;
  return getTokenViaDialog();
}

// ── Graph API ─────────────────────────────────────────────────────────────────
async function writeGraphExtension(token, restMessageId, data) {
  const baseUrl = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions`;
  const body = JSON.stringify({
    "@odata.type": "microsoft.graph.openTypeExtension",
    extensionName: EXTENSION_NAME,
    ...data
  });
  const headers = {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json"
  };

  let res = await fetch(baseUrl, { method: "POST", headers, body });

  if (res.status === 409) {
    res = await fetch(`${baseUrl}/${EXTENSION_NAME}`, { method: "PATCH", headers, body });
    if (res.status !== 204) throw new Error(`Extension update failed: ${res.status}`);
    return;
  }

  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new Error(`Graph API ${res.status}: ${detail}`);
  }
}

// ── Outlook (Office.js) path ──────────────────────────────────────────────────
if (typeof Office !== "undefined") {
  Office.onReady(() => {
    initForm();

    // Read email properties first — isolated from MSAL so auth errors can't block this
    try {
      const item = Office.context.mailbox.item;
      document.getElementById("from").value        = item.from ? item.from.emailAddress : "";
      document.getElementById("subject").value     = item.subject || "";
      document.getElementById("f-reference").value = item.subject || "";
    } catch (e) {
      console.warn("Could not read email properties:", e);
    }

    try { initMsal(); } catch (e) { console.warn("MSAL init failed:", e); }

    document.getElementById("save-btn").addEventListener("click", async () => {
      const btn = document.getElementById("save-btn");
      btn.disabled = true;
      showStatus("Authenticating…", false);
      try {
        const token = await getToken();
        showStatus("Saving…", false);
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        await writeGraphExtension(token, restId, buildOutput());
        showStatus("Saved to message extension.", false);
      } catch (e) {
        showStatus(e.message || "Save failed.", true);
        console.error(e);
      } finally {
        btn.disabled = false;
      }
    });
  });
} else {
  // ── Browser preview / testbed path ───────────────────────────────────────────
  document.addEventListener("DOMContentLoaded", () => {
    initForm();
    document.getElementById("from").value        = "sender@example.com";
    document.getElementById("subject").value     = "[PREVIEW] Sample Email Subject";
    document.getElementById("f-reference").value = "[PREVIEW] Sample Email Subject";

    const btn = document.getElementById("save-btn");
    if (btn) {
      btn.addEventListener("click", () => {
        const out = document.getElementById("preview-output");
        if (out) {
          out.textContent = JSON.stringify(buildOutput(), null, 2);
          out.style.display = "block";
        }
      });
    }
  });
}
