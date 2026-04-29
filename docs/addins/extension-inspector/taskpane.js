/* global Office, getCachedToken, cacheToken */

const CLIENT_ID  = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const DIALOG_URL = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";

const META = new Set(["@odata.type", "@odata.context", "@odata.etag", "id"]);

const LOG = (...a) => console.log("[ExtInspector]", ...a);

// ── Auth ──────────────────────────────────────────────────────────────────────
function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    LOG("opening auth dialog:", DIALOG_URL);
    Office.context.ui.displayDialogAsync(
      DIALOG_URL,
      { height: 60, width: 30, promptBeforeOpen: false },
      (result) => {
        LOG("displayDialogAsync callback, status:", result.status);
        if (result.status === Office.AsyncResultStatus.Failed) {
          LOG("dialog open failed:", result.error);
          reject(new Error(`Dialog failed: ${result.error.message}`));
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
          LOG("message from dialog:", msg.message);
          dialog.close();
          try {
            const payload = JSON.parse(msg.message);
            if (payload.error) {
              LOG("auth error in payload:", payload.error);
              reject(new Error(payload.error));
            } else {
              LOG("token received, expiresIn:", payload.expiresIn);
              cacheToken(payload.accessToken, payload.expiresIn ?? 3600);
              resolve(payload.accessToken);
            }
          } catch (e) {
            LOG("failed to parse dialog message:", e);
            reject(new Error("Invalid message from auth dialog"));
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
          LOG("dialog event:", evt.error);
          if (evt.error === 12006) reject(new Error("Sign-in cancelled."));
        });
      }
    );
  });
}

async function getToken() {
  const cached = getCachedToken();
  LOG("cached token present:", !!cached);
  if (cached) return cached;
  return getTokenViaDialog();
}

// ── Graph ────────────────────────────────────────────────────────────────────
async function fetchAllExtensions(token, restMessageId) {
  // The /extensions navigation property is the correct endpoint for listing all
  // extensions. It works for work/school accounts. Personal MSA accounts return
  // 405 — in that case we surface the manual lookup UI instead.
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions`;
  LOG("fetching all extensions:", url);
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  LOG("response status:", res.status);
  if (res.ok) {
    const json = await res.json();
    LOG("extensions:", json.value);
    return { extensions: json.value ?? [], msaLimit: false };
  }
  if (res.status === 405) {
    LOG("405 — personal MSA account; nav property not supported");
    return { extensions: [], msaLimit: true };
  }
  const text = await res.text();
  LOG("response body:", text);
  throw new Error(`Graph ${res.status}: ${text}`);
}

async function fetchExtensionByName(token, restMessageId, name) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions/${encodeURIComponent(name)}`;
  LOG("fetching by name:", url);
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  LOG("response status:", res.status);
  if (res.status === 404) return null;
  const text = await res.text();
  LOG("response body:", text);
  if (!res.ok) throw new Error(`Graph ${res.status}: ${text}`);
  return JSON.parse(text);
}

// ── Rendering ─────────────────────────────────────────────────────────────────
function renderValue(val, depth) {
  if (val === null) return document.createTextNode("null");
  if (Array.isArray(val)) return renderArray(val, depth);
  if (typeof val === "object") return renderObject(val, depth);
  return document.createTextNode(String(val));
}

function renderObject(obj, depth) {
  const keys = Object.keys(obj).filter(k => !META.has(k));
  if (keys.length === 0) return document.createTextNode("{}");
  const wrap = document.createElement("div");
  wrap.className = depth > 0 ? "nested" : "";
  const table = document.createElement("table");
  table.className = "kv-table";
  keys.forEach(k => {
    const tr = document.createElement("tr");
    const tdKey = document.createElement("td");
    tdKey.className = "key";
    tdKey.textContent = k;
    const tdVal = document.createElement("td");
    tdVal.className = "val";
    tdVal.appendChild(renderValue(obj[k], depth + 1));
    tr.appendChild(tdKey);
    tr.appendChild(tdVal);
    table.appendChild(tr);
  });
  wrap.appendChild(table);
  return wrap;
}

function renderArray(arr, depth) {
  if (arr.length === 0) return document.createTextNode("[]");
  const wrap = document.createElement("div");
  arr.forEach((item, i) => {
    const itemWrap = document.createElement("div");
    itemWrap.className = "array-item";
    if (arr.length > 1) {
      const idx = document.createElement("div");
      idx.className = "array-index";
      idx.textContent = `[${i}]`;
      itemWrap.appendChild(idx);
    }
    itemWrap.appendChild(renderValue(item, depth));
    wrap.appendChild(itemWrap);
  });
  return wrap;
}

function renderExtension(ext) {
  const card = document.createElement("div");
  card.className = "ext-card";
  const header = document.createElement("div");
  header.className = "ext-card-header";
  const rawId = ext.id ?? "(unknown extension)";
  header.textContent = rawId.replace(/^Microsoft\.OutlookServices\.OpenTypeExtension\./, "");
  card.appendChild(header);
  const body = document.createElement("div");
  body.className = "ext-card-body";
  body.appendChild(renderObject(ext, 0));
  card.appendChild(body);
  return card;
}

function setStatus(msg, isError) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = isError ? "error" : "";
}

function appendExtension(ext) {
  document.getElementById("extensions-container").appendChild(renderExtension(ext));
}

// ── Entry point ───────────────────────────────────────────────────────────────
Office.onReady(async () => {
  LOG("Office.onReady fired");
  try {
    const item = Office.context.mailbox.item;
    document.getElementById("ei-from").textContent    = item.from?.emailAddress ?? "—";
    document.getElementById("ei-subject").textContent = item.subject ?? "—";

    LOG("ewsId (raw):", item.itemId);
    const restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    LOG("restId:", restId);

    setStatus("Authenticating…");
    const token = await getToken();

    setStatus("Loading extensions…");
    const { extensions, msaLimit } = await fetchAllExtensions(token, restId);

    if (msaLimit) {
      // Personal MSA accounts: surface the manual lookup UI
      setStatus("");
      document.getElementById("msa-fallback").style.display = "block";
      document.getElementById("lookup-btn").addEventListener("click", async () => {
        const name = document.getElementById("ext-name-input").value.trim();
        if (!name) return;
        setStatus("Looking up…");
        document.getElementById("extensions-container").innerHTML = "";
        try {
          const ext = await fetchExtensionByName(token, restId, name);
          setStatus("");
          if (!ext) {
            setStatus(`No extension found: ${name}`, true);
          } else {
            appendExtension(ext);
          }
        } catch (e) {
          setStatus(e.message || "Lookup failed.", true);
          LOG("lookup error:", e);
        }
      });
      return;
    }

    if (extensions.length === 0) {
      setStatus("");
      const empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = "No extensions found on this message.";
      document.getElementById("extensions-container").appendChild(empty);
      return;
    }

    setStatus("");
    extensions.forEach(appendExtension);
  } catch (e) {
    setStatus(e.message || "Failed to load extensions.", true);
    LOG("error:", e);
  }
});
