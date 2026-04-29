/* global Office, getCachedToken, cacheToken, readGraphExtension, writeGraphExtension */

// ── Configuration ─────────────────────────────────────────────────────────────
const CLIENT_ID      = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const EXTENSION_NAME = "net.cprima.rpapub.EnfoldedOrichalcum";
const DIALOG_URL     = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";
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

function populateFormFromExtension(ext) {
  const map = {
    reference: "f-reference", category: "f-category", priority: "f-priority",
    owner: "f-owner", notes: "f-notes", country: "f-country",
    year: "f-year", quarter: "f-quarter"
  };
  Object.entries(map).forEach(([key, id]) => {
    if (ext[key] != null && ext[key] !== "") document.getElementById(id).value = ext[key];
  });
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

// ── Auth via Office Dialog ────────────────────────────────────────────────────
// MSAL is NOT loaded in the taskpane (blocked by Edge Tracking Prevention).
// Auth runs entirely inside auth-dialog.html, a real browser window, which
// sends the token back via Office.context.ui.messageParent.
function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      DIALOG_URL,
      { height: 60, width: 30, promptBeforeOpen: true },
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
            else {
              cacheToken(payload.accessToken, payload.expiresIn ?? 3600);
              resolve(payload.accessToken);
            }
          } catch (e) {
            reject(new Error("Invalid message from auth dialog"));
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
          // 12006 = user closed dialog; ignore during redirect cycle (fires on navigation away)
          if (evt.error === 12006) reject(new Error("Sign-in cancelled."));
        });
      }
    );
  });
}

async function getToken() {
  const cached = getCachedToken();
  if (cached) return cached;
  return getTokenViaDialog();
}

const FORM_STORAGE_KEY = "triage_pending_form";

function saveFormToStorage() {
  localStorage.setItem(FORM_STORAGE_KEY, JSON.stringify(buildOutput()));
}

function restoreFormFromStorage() {
  const raw = localStorage.getItem(FORM_STORAGE_KEY);
  if (!raw) return false;
  try {
    const d = JSON.parse(raw);
    ["f-reference","f-category","f-priority","f-owner","f-notes"].forEach(id => {
      if (d[id.replace("f-","")] !== undefined)
        document.getElementById(id).value = d[id.replace("f-","")];
    });
    if (d.country) document.getElementById("f-country").value = d.country;
    if (d.year)    document.getElementById("f-year").value    = d.year;
    if (d.quarter) document.getElementById("f-quarter").value = d.quarter;
    return true;
  } catch (_) { return false; }
}

// ── Outlook (Office.js) path ──────────────────────────────────────────────────
if (typeof Office !== "undefined") {
  Office.onReady(async () => {
    initForm();

    const item = Office.context.mailbox.item;
    try {
      document.getElementById("from").value        = item.from ? item.from.emailAddress : "";
      document.getElementById("subject").value     = item.subject || "";
      document.getElementById("f-reference").value = item.subject || "";
      document.getElementById("message-id").value  = item.itemId || "";
    } catch (e) {
      console.warn("Could not read email properties:", e);
    }

    // Restore any form data saved before the auth redirect reloaded the task pane.
    // If there was pending data, skip the Graph read (user was mid-edit).
    const hadPendingForm = restoreFormFromStorage();

    if (!hadPendingForm) {
      const cached = getCachedToken();
      if (cached) {
        try {
          const restId = Office.context.mailbox.convertToRestId(
            item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
          );
          const ext = await readGraphExtension(cached, restId, EXTENSION_NAME);
          if (ext) populateFormFromExtension(ext);
        } catch (_) { /* silent — form stays at defaults */ }
      }
    }

    document.getElementById("save-btn").addEventListener("click", async () => {
      const btn = document.getElementById("save-btn");
      btn.disabled = true;
      showStatus("Authenticating…", false);
      // Persist form before auth dialog (redirect may reload task pane)
      saveFormToStorage();
      try {
        const token = await getToken();
        showStatus("Saving…", false);
        const data = buildOutput();
        localStorage.removeItem(FORM_STORAGE_KEY);
        const restId = Office.context.mailbox.convertToRestId(
          Office.context.mailbox.item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        await writeGraphExtension(token, restId, EXTENSION_NAME, data);
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
