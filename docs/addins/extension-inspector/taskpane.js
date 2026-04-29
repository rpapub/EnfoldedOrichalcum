/* global Office, getCachedToken, cacheToken */

const CLIENT_ID  = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const DIALOG_URL = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";

// Graph metadata fields to suppress in the rendered output
const META = new Set(["@odata.type", "@odata.context", "@odata.etag", "id"]);

// ── Auth (same pattern as embed-data) ────────────────────────────────────────
function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      DIALOG_URL,
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

// ── Graph ────────────────────────────────────────────────────────────────────
async function fetchExtensions(token, restMessageId) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text().catch(() => "")}`);
  const json = await res.json();
  return json.value ?? [];
}

// ── Rendering ────────────────────────────────────────────────────────────────
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
  header.textContent = ext.id ?? "(unknown extension)";
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

// ── Entry point ───────────────────────────────────────────────────────────────
Office.onReady(async () => {
  try {
    const item = Office.context.mailbox.item;
    document.getElementById("ei-from").textContent    = item.from?.emailAddress ?? "—";
    document.getElementById("ei-subject").textContent = item.subject ?? "—";

    const restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    setStatus("Authenticating…");
    const token = await getToken();

    setStatus("Loading extensions…");
    const extensions = await fetchExtensions(token, restId);

    const container = document.getElementById("extensions-container");
    if (extensions.length === 0) {
      setStatus("");
      const empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = "No extensions found on this message.";
      container.appendChild(empty);
      return;
    }

    setStatus("");
    extensions.forEach(ext => container.appendChild(renderExtension(ext)));
  } catch (e) {
    setStatus(e.message || "Failed to load extensions.", true);
    console.error(e);
  }
});
