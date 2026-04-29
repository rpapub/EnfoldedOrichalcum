/* global Office, getCachedToken, cacheToken, clearCachedToken, readGraphExtension, writeGraphExtension */

const CLIENT_ID        = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const EXTENSION_NAME   = "net.cprima.rpapub.CPMForge.M365.triage";
const DIALOG_URL       = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";
const FORM_STORAGE_KEY = "mail_triage_pending_form";
const SCHEMA_URL       = "schema.json";
const GRAPH_META       = new Set(["@odata.type", "@odata.context", "@odata.etag", "id"]);

let currentStateVersion = 0;
let triageSchema = null;

// ── Schema ────────────────────────────────────────────────────────────────────
async function loadSchema() {
  try {
    const r = await fetch(SCHEMA_URL);
    triageSchema = await r.json();
  } catch (_) {
    triageSchema = { properties: {} };
  }
}

function evidenceProp(key) {
  try { return triageSchema.properties.evidence.properties[key] || null; }
  catch (_) { return null; }
}

function evidenceProps() {
  try { return triageSchema.properties.evidence.properties || {}; }
  catch (_) { return {}; }
}

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
  populateSelect("f-triage-result", [
    ["no_action","no_action"], ["work_required","work_required"],
    ["information_only","information_only"], ["automated_candidate","automated_candidate"],
    ["escalation_needed","escalation_needed"]
  ]);
  populateSelect("f-triage-confidence", [
    ["low","low"], ["medium","medium"], ["high","high"]
  ]);
}

function showStatus(msg, isError) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status " + (isError ? "error" : "ok");
  if (!isError) setTimeout(() => { el.textContent = ""; el.className = "status"; }, 4000);
}

// ── Evidence fields ───────────────────────────────────────────────────────────
function shownEvidenceKeys() {
  const keys = new Set();
  document.querySelectorAll("#evidence-fields [data-ev-key]").forEach(el => keys.add(el.dataset.evKey));
  return keys;
}

function appendEvidenceField(key, val, customType) {
  const container = document.getElementById("evidence-fields");
  if (container.querySelector(`[data-ev-key="${CSS.escape(key)}"]`)) return;

  const prop = evidenceProp(key);
  const type = prop ? prop.type : (customType || "string");
  const title = (prop && prop.title) ? prop.title : key;

  const div = document.createElement("div");
  div.className = "field ev-field";
  div.dataset.evKey = key;

  const lbl = document.createElement("label");
  lbl.textContent = title;

  const rm = document.createElement("button");
  rm.type = "button";
  rm.className = "ev-remove";
  rm.textContent = "x";
  rm.title = "Remove";
  rm.addEventListener("click", () => { div.remove(); refreshAddRow(); });

  let input;
  if (prop && prop.enum) {
    input = document.createElement("select");
    input.dataset.evKey = key;
    const blank = document.createElement("option");
    blank.value = ""; blank.textContent = "-- not set --";
    input.appendChild(blank);
    prop.enum.forEach(v => {
      const opt = document.createElement("option");
      opt.value = v; opt.textContent = v;
      input.appendChild(opt);
    });
    if (val != null && val !== "") input.value = String(val);
  } else if (type === "boolean") {
    const row = document.createElement("div");
    row.className = "ev-bool-row";
    input = document.createElement("input");
    input.type = "checkbox";
    input.dataset.evKey = key;
    if (val === true || val === "true") input.checked = true;
    row.appendChild(input);
    div.appendChild(lbl);
    div.appendChild(row);
    div.appendChild(rm);
    container.insertBefore(div, document.getElementById("evidence-add-row"));
    refreshAddRow();
    return;
  } else if (type === "integer") {
    input = document.createElement("input");
    input.type = "number";
    input.step = "1";
    input.dataset.evKey = key;
    if (val != null) input.value = String(val);
  } else {
    input = document.createElement("input");
    input.type = "text";
    input.dataset.evKey = key;
    if (val != null) input.value = String(val);
  }

  div.appendChild(lbl);
  div.appendChild(input);
  div.appendChild(rm);
  container.insertBefore(div, document.getElementById("evidence-add-row"));
  refreshAddRow();
}

function addDefaultEvidenceFields() {
  const props = evidenceProps();
  Object.entries(props).forEach(([key, prop]) => {
    if (prop["x-default"]) appendEvidenceField(key, null);
  });
}

// ── Attachment detection ──────────────────────────────────────────────────────
function detectAttachmentProfile(attachments) {
  if (!attachments || attachments.length === 0) return "none";
  const files = attachments.filter(a => !a.isInline);
  if (files.length === 0) return "none";
  const names = files.map(a => (a.name || "").toLowerCase());
  const hasPdf   = names.some(n => n.endsWith(".pdf"));
  const hasExcel = names.some(n => /\.(xlsx?|xlsm|xlsb)$/.test(n));
  const hasOther = names.some(n => !/\.(pdf|xlsx?|xlsm|xlsb)$/.test(n));
  if (hasPdf && hasExcel && !hasOther) return "pdf_and_excel";
  if (hasPdf && !hasExcel && !hasOther) return "pdf_only";
  if (!hasPdf && hasExcel && !hasOther) return "excel_only";
  if (hasOther && (hasPdf || hasExcel)) return "mixed";
  return "other";
}

function applyDetectedAttachmentProfile(item) {
  if (document.querySelector('#evidence-fields [data-ev-key="attachmentProfile"]')) return;
  const profile = detectAttachmentProfile(item.attachments);
  appendEvidenceField("attachmentProfile", profile);
}

function refreshAddRow() {
  const row = document.getElementById("evidence-add-row");
  row.innerHTML = "";

  const shown = shownEvidenceKeys();
  const props = evidenceProps();
  const available = Object.keys(props).filter(k => !shown.has(k));

  const fieldSel = document.createElement("select");
  fieldSel.className = "ev-add-select";
  const ph = document.createElement("option");
  ph.value = ""; ph.textContent = "Add field...";
  fieldSel.appendChild(ph);
  available.forEach(key => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = props[key].title || key;
    fieldSel.appendChild(opt);
  });
  const custOpt = document.createElement("option");
  custOpt.value = "__custom__"; custOpt.textContent = "Custom...";
  fieldSel.appendChild(custOpt);

  const customKeyInput = document.createElement("input");
  customKeyInput.type = "text";
  customKeyInput.placeholder = "key name";
  customKeyInput.className = "ev-custom-key";
  customKeyInput.style.display = "none";

  const typeSel = document.createElement("select");
  typeSel.className = "ev-type-select";
  typeSel.style.display = "none";
  [["string","text"], ["integer","integer"], ["boolean","boolean"]].forEach(([v,l]) => {
    const opt = document.createElement("option");
    opt.value = v; opt.textContent = l;
    typeSel.appendChild(opt);
  });

  fieldSel.addEventListener("change", () => {
    const isCustom = fieldSel.value === "__custom__";
    customKeyInput.style.display = isCustom ? "" : "none";
    typeSel.style.display = isCustom ? "" : "none";
  });

  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.className = "ev-add-btn";
  addBtn.textContent = "+";
  addBtn.addEventListener("click", () => {
    const v = fieldSel.value;
    if (!v) return;
    if (v === "__custom__") {
      const k = customKeyInput.value.trim();
      if (!k) { customKeyInput.focus(); return; }
      appendEvidenceField(k, null, typeSel.value);
    } else {
      appendEvidenceField(v, null);
    }
    fieldSel.value = "";
    customKeyInput.value = "";
    customKeyInput.style.display = "none";
    typeSel.style.display = "none";
  });

  row.appendChild(fieldSel);
  row.appendChild(customKeyInput);
  row.appendChild(typeSel);
  row.appendChild(addBtn);
}

function collectEvidence() {
  const obj = {};
  document.querySelectorAll("#evidence-fields [data-ev-key]").forEach(input => {
    const key = input.dataset.evKey;
    if (input.type === "checkbox") {
      obj[key] = input.checked;
    } else if (input.type === "number") {
      if (input.value.trim() !== "") obj[key] = Number(input.value);
    } else {
      if (input.value.trim() !== "") obj[key] = input.value.trim();
    }
  });
  return obj;
}

// ── Output ────────────────────────────────────────────────────────────────────
function buildOutput() {
  const evidenceObj = collectEvidence();
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
  document.getElementById("f-state-version").value = ext.stateVersion || "1";

  if (ext.triage) {
    if (ext.triage.result)          document.getElementById("f-triage-result").value     = ext.triage.result;
    if (ext.triage.confidence)      document.getElementById("f-triage-confidence").value = ext.triage.confidence;
    if (ext.triage.summary != null) document.getElementById("f-triage-summary").value    = ext.triage.summary;
  }

  if (ext.evidence) {
    Object.entries(ext.evidence).forEach(([key, val]) => {
      if (GRAPH_META.has(key)) return;
      appendEvidenceField(key, val);
    });
  }
}

// ── Form persistence ──────────────────────────────────────────────────────────
function saveFormToStorage() {
  localStorage.setItem(FORM_STORAGE_KEY, JSON.stringify({
    caseId:           document.getElementById("f-case-id").value,
    caseStatus:       document.getElementById("f-case-status").value,
    triageResult:     document.getElementById("f-triage-result").value,
    triageConfidence: document.getElementById("f-triage-confidence").value,
    triageSummary:    document.getElementById("f-triage-summary").value,
    evidence:         collectEvidence(),
    stateVersion:     currentStateVersion,
  }));
}

function restoreFormFromStorage() {
  const raw = localStorage.getItem(FORM_STORAGE_KEY);
  if (!raw) return false;
  try {
    const d = JSON.parse(raw);
    if (d.caseId)           document.getElementById("f-case-id").value           = d.caseId;
    if (d.caseStatus)       document.getElementById("f-case-status").value       = d.caseStatus;
    if (d.triageResult)     document.getElementById("f-triage-result").value     = d.triageResult;
    if (d.triageConfidence) document.getElementById("f-triage-confidence").value = d.triageConfidence;
    if (d.triageSummary != null) document.getElementById("f-triage-summary").value = d.triageSummary;
    if (d.evidence) {
      Object.entries(d.evidence).forEach(([key, val]) => appendEvidenceField(key, val));
    }
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
  const item = Office.context.mailbox.item;
  const cached = getCachedToken();
  if (!cached) {
    document.getElementById("f-case-id").value       = generateCaseId();
    document.getElementById("f-case-status").value   = "new";
    document.getElementById("f-state-version").value = "1";
    applyDetectedAttachmentProfile(item);
    return;
  }
  const statusEl = document.getElementById("status");
  statusEl.textContent = "Loading...";
  statusEl.className = "status";
  try {
    const restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    const ext = await readGraphExtension(cached, restId, EXTENSION_NAME);
    if (ext) {
      populateFormFromExtension(ext);
      // fill attachmentProfile only if the stored extension didn't include it
      applyDetectedAttachmentProfile(item);
    } else {
      document.getElementById("f-case-id").value          = generateCaseId();
      document.getElementById("f-case-status").value      = "new";
      document.getElementById("f-state-version").value    = "1";
      currentStateVersion = 0;
      applyDetectedAttachmentProfile(item);
    }
  } catch (e) {
    if (e.status === 401) clearCachedToken();
  } finally {
    if (statusEl.textContent === "Loading...") {
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
    await loadSchema();
    refreshAddRow();

    const item = Office.context.mailbox.item;
    try {
      document.getElementById("ei-from").textContent    = item.from ? item.from.emailAddress : "--";
      document.getElementById("ei-subject").textContent = item.subject || "--";
    } catch (e) {
      console.warn("Could not read email properties:", e);
    }

    document.getElementById("save-btn").addEventListener("click", async () => {
      const btn = document.getElementById("save-btn");
      btn.disabled = true;
      showStatus("Authenticating...", false);
      saveFormToStorage();
      try {
        const token = await getToken();
        showStatus("Saving...", false);
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
          showStatus("Session expired -- please sign in again and retry.", true);
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
  document.addEventListener("DOMContentLoaded", async () => {
    initSelects();
    await loadSchema();
    refreshAddRow();
    document.getElementById("ei-from").textContent    = "sender@example.com";
    document.getElementById("ei-subject").textContent = "[PREVIEW] Sample Email Subject";
    document.getElementById("f-case-id").value        = generateCaseId();
    document.getElementById("f-case-status").value    = "new";
    addDefaultEvidenceFields();

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
