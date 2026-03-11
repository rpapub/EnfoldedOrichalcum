/* global Office */

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
  const month = new Date().getMonth(); // 0-based
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
  setTimeout(() => { el.textContent = ""; el.className = "status"; }, 3000);
}

// ── Outlook (Office.js) path ──────────────────────────────────────────────────
if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      initForm();
      const item = Office.context.mailbox.item;
      document.getElementById("from").value    = item.from ? item.from.emailAddress : "";
      document.getElementById("subject").value  = item.subject || "";
      document.getElementById("f-reference").value = item.subject || "";
      document.getElementById("copy-btn").addEventListener("click", copyToClipboard);
    }
  });
} else {
  // ── Browser preview / testbed path ───────────────────────────────────────────
  document.addEventListener("DOMContentLoaded", () => {
    initForm();
    document.getElementById("from").value     = "sender@example.com";
    document.getElementById("subject").value   = "[PREVIEW] Sample Email Subject";
    document.getElementById("f-reference").value = "[PREVIEW] Sample Email Subject";
    const btn = document.getElementById("copy-btn");
    if (btn) btn.addEventListener("click", copyToClipboard);
  });
}

function copyToClipboard() {
  const data = buildOutput();
  const text = JSON.stringify(data, null, 2);
  navigator.clipboard.writeText(text).then(
    () => showStatus("Copied to clipboard."),
    () => showStatus("Clipboard unavailable — see console.", true)
  );
}
