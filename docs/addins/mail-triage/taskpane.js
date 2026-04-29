/* global Office, getCachedToken, cacheToken, clearCachedToken, readGraphExtension, writeGraphExtension */

const CLIENT_ID        = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const EXTENSION_NAME   = "net.cprima.rpapub.CPMForge.M365.triage";
const DIALOG_URL       = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";
const FORM_STORAGE_KEY = "mail_triage_pending_form";

let currentStateVersion = 0;

// ── Helpers ───────────────────────────────────────────────────────────────────
function generateCaseId() {
  return "C" + Date.now().toString(36).toUpperCase();
}

function populateSelect(id, options) {
  const sel = document.getElementById(id);
  sel.innerHTML = "";
  options.forEach(([val, label]) => {
    const opt = document.createElement("option");
    opt.value = val;
    opt.textContent = label || val;
    sel.appendChild(opt);
  });
}

function initSelects() {
  populateSelect("f-case-status", [
    ["new","new"], ["triaged","triaged"], ["in_progress","in_progress"],
    ["resolved","resolved"], ["closed","closed"]
  ]);
  // no_action and low are first → selected by default for untriaged messages
  populateSelect("f-triage-result", [
    ["no_action","no_action"], ["work_required","work_required"],
    ["information_only","information_only"], ["automated_candidate","automated_candidate"],
    ["escalation_needed","escalation_needed"]
  ]);
  populateSelect("f-triage-confidence", [
    ["low","low"], ["medium","medium"], ["high","high"]
  ]);
  populateSelect("f-attachment-profile", [
    ["","— not set —"],
    ["none","none"], ["pdf_only","pdf_only"], ["excel_only","excel_only"],
    ["pdf_and_excel","pdf_and_excel"], ["mixed","mixed"], ["other","other"]
  ]);
}

function showStatus(msg, isError) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status " + (isError ? "error" : "ok");
  if (!isError) setTimeout(() => { el.textContent = ""; el.className = "status"; }, 4000);
}

// ── Output ────────────────────────────────────────────────────────────────────
function buildOutput() {
  const hasBodyEl         = document.getElementById("f-has-body-instruction");
  const attachmentProfile = document.getElementById("f-attachment-profile").value;
  const evidenceSummary   = document.getElementById("f-evidence-summary").value.trim();

  const evidenceObj = {};
  if (hasBodyEl.dataset.set)  evidenceObj.hasBodyInstruction = hasBodyEl.checked;
  if (attachmentProfile)      evidenceObj.attachmentProfile  = attachmentProfile;
  if (evidenceSummary)        evidenceObj.evidenceSummary    = evidenceSummary;

  const out = {
    caseId:       document.getElementById("f-case-id").value.trim(),
    caseStatus:   document.getElementById("f-case-status").value,
    stateVersion: String(currentStateVersion),
    triage: {
      result:     document.getElementById("f-triage-result").value,
      confidence: document.getElementById("f-triage-confidence").value,
      summary:    document.getElementById("f-triage-summary").value.trim(),
    },
  };
  if (Object.keys(evidenceObj).length > 0) out.evidence = evidenceObj;
  return out;
}

// ── Form / extension ──────────────────────────────────────────────────────────
function populateFormFromExtension(ext) {
  if (ext.caseId)     document.getElementById("f-case-id").value     = ext.caseId;
  if (ext.caseStatus) document.getElementById("f-case-status").value = ext.caseStatus;
  currentStateVersion = parseInt(ext.stateVersion || "0", 10);
  document.getElementById("f-state-version").value = ext.stateVersion || "0";

  if (ext.triage) {
    if (ext.triage.result)          document.getElementById("f-triage-result").value     = ext.triage.result;
    if (ext.triage.confidence)      document.getElementById("f-triage-confidence").value = ext.triage.confidence;
    if (ext.triage.summary != null) document.getElementById("f-triage-summary").value    = ext.triage.summary;
  }

  // evidence is optional; only set fields that are present in the extension
  const e = ext.evidence || {};
  const hasBodyEl = document.getElementById("f-has-body-instruction");
  if (e.hasBodyInstruction != null) {
    hasBodyEl.checked = e.hasBodyInstruction;
    hasBodyEl.dataset.set = "1";
  }
  if (e.attachmentProfile != null) {
    document.getElementById("f-attachment-profile").value = e.attachmentProfile;
  }
  if (e.evidenceSummary != null) {
    document.getElementById("f-evidence-summary").value = e.evidenceSummary;
  }
}

// ── Form persistence ──────────────────────────────────────────────────────────
function saveFormToStorage() {
  const hasBodyEl = document.getElementById("f-has-body-instruction");
  localStorage.setItem(FORM_STORAGE_KEY, JSON.stringify({
    caseId:             document.getElementById("f-case-id").value,
    caseStatus:         document.getElementById("f-case-status").value,
    triageResult:       document.getElementById("f-triage-result").value,
    triageConfidence:   document.getElementById("f-triage-confidence").value,
    triageSummary:      document.getElementById("f-triage-summary").value,
    hasBodyInstruction: hasBodyEl.checked,
    hasBodySet:         hasBodyEl.dataset.set || "",
    attachmentProfile:  document.getElementById("f-attachment-profile").value,
    evidenceSummary:    document.getElementById("f-evidence-summary").value,
    stateVersion:       currentStateVersion,
  }));
}

function restoreFormFromStorage() {
  const raw = localStorage.getItem(FORM_STORAGE_KEY);
  if (!raw) return false;
  try {
    const d = JSON.parse(raw);
    if (d.caseId)            document.getElementById("f-case-id").value           = d.caseId;
    if (d.caseStatus)        document.getElementById("f-case-status").value       = d.caseStatus;
    if (d.triageResult)      document.getElementById("f-triage-result").value     = d.triageResult;
    if (d.triageConfidence)  document.getElementById("f-triage-confidence").value = d.triageConfidence;
    if (d.triageSummary != null) document.getElementById("f-triage-summary").value = d.triageSummary;
    const hasBodyEl = document.getElementById("f-has-body-instruction");
    hasBodyEl.checked = !!d.hasBodyInstruction;
    if (d.hasBodySet)        hasBodyEl.dataset.set = "1";
    if (d.attachmentProfile) document.getElementById("f-attachment-profile").value = d.attachmentProfile;
    if (d.evidenceSummary != null) document.getElementById("f-evidence-summary").value = d.evidenceSummary;
    if (d.stateVersion != null) currentStateVersion = parseInt(d.stateVersion, 10);
    return true;
  } catch (_) { return false; }
}

// ── Auth via Office Dialog ────────────────────────────────────────────────────
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

// ── Graph prefill ─────────────────────────────────────────────────────────────
async function prefillFromExtension() {
  const cached = getCachedToken();
  if (!cached) {
    document.getElementById("f-case-id").value     = generateCaseId();
    document.getElementById("f-case-status").value = "new";
    return;
  }
  const statusEl = document.getElementById("status");
  statusEl.textContent = "Loading…";
  statusEl.className = "status";
  try {
    const restId = Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    const ext = await readGraphExtension(cached, restId, EXTENSION_NAME);
    if (ext) {
      populateFormFromExtension(ext);
    } else {
      document.getElementById("f-case-id").value     = generateCaseId();
      document.getElementById("f-case-status").value = "new";
      currentStateVersion = 0;
    }
  } catch (e) {
    if (e.status === 401) clearCachedToken();
  } finally {
    if (statusEl.textContent === "Loading…") {
      statusEl.textContent = "";
      statusEl.className = "status";
    }
  }
}

// ── Outlook (Office.js) path ──────────────────────────────────────────────────
if (typeof Office !== "undefined") {
  Office.onReady(async () => {
    const taglines = [
      "triage once. route always.",
      "status. confidence. action.",
      "from inbox to workflow",
      "classify. assign. resolve.",
      "structured triage for structured teams",
      "less noise. more signal.",
      "one click closer to resolved."
    ];
    document.getElementById("brand-bar").textContent =
      taglines[new Date().getHours() % taglines.length];

    initSelects();

    const item = Office.context.mailbox.item;
    try {
      document.getElementById("ei-from").textContent    = item.from ? item.from.emailAddress : "—";
      document.getElementById("ei-subject").textContent = item.subject || "—";
    } catch (e) {
      console.warn("Could not read email properties:", e);
    }

    document.getElementById("save-btn").addEventListener("click", async () => {
      const btn = document.getElementById("save-btn");
      btn.disabled = true;
      showStatus("Authenticating…", false);
      saveFormToStorage();
      try {
        const token = await getToken();
        showStatus("Saving…", false);
        currentStateVersion++;
        const data = buildOutput();
        localStorage.removeItem(FORM_STORAGE_KEY);
        const restId = Office.context.mailbox.convertToRestId(
          item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        await writeGraphExtension(token, restId, EXTENSION_NAME, data);
        document.getElementById("f-state-version").value = String(currentStateVersion);
        showStatus("Saved.", false);
      } catch (e) {
        currentStateVersion--;
        if (e.status === 401) {
          clearCachedToken();
          showStatus("Session expired — please sign in again and retry.", true);
        } else {
          showStatus(e.message || "Save failed.", true);
        }
        console.error(e);
      } finally {
        btn.disabled = false;
      }
    });

    document.getElementById("signout-btn").addEventListener("click", ev => {
      ev.preventDefault();
      clearCachedToken();
      location.reload();
    });

    const hadPendingForm = restoreFormFromStorage();
    if (!hadPendingForm) await prefillFromExtension();
  });
} else {
  // ── Browser preview / testbed path ───────────────────────────────────────────
  document.addEventListener("DOMContentLoaded", () => {
    initSelects();
    document.getElementById("ei-from").textContent    = "sender@example.com";
    document.getElementById("ei-subject").textContent = "[PREVIEW] Sample Email Subject";
    document.getElementById("f-case-id").value        = generateCaseId();
    document.getElementById("f-case-status").value    = "new";

    document.getElementById("save-btn").addEventListener("click", () => {
      currentStateVersion++;
      const out = document.getElementById("preview-output");
      if (out) {
        out.textContent = JSON.stringify(buildOutput(), null, 2);
        out.style.display = "block";
      }
    });
  });
}
