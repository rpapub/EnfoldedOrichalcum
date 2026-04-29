/* global Office, getCachedToken, cacheToken, clearCachedToken, graphError, deleteGraphExtension */

const CLIENT_ID       = "f28629c6-2f87-4afc-a6ff-1cbbd50166af";
const DIALOG_URL      = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";
const DEFAULT_EXT     = "net.cprima.rpapub.CPMForge.M365.extension";
const LAST_EXT_KEY    = "ext_inspector_last_name";

const META = new Set(["@odata.type", "@odata.context", "@odata.etag", "id"]);

// ── Auth ──────────────────────────────────────────────────────────────────────
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
            if (payload.error) {
              reject(new Error(payload.error));
            } else {
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
async function fetchAllExtensions(token, restMessageId) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.ok) {
    const json = await res.json();
    return json.value ?? [];
  }
  const text = await res.text().catch(() => "");
  throw graphError(res.status, text);
}

async function fetchExtensionByName(token, restMessageId, name) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions/${encodeURIComponent(name)}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  const text = await res.text();
  if (!res.ok) throw graphError(res.status, text);
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
    const tdKey = document.createElement("td"); tdKey.className = "key"; tdKey.textContent = k;
    const tdVal = document.createElement("td"); tdVal.className = "val";
    tdVal.appendChild(renderValue(obj[k], depth + 1));
    tr.appendChild(tdKey); tr.appendChild(tdVal);
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

function renderExtension(ext, onDelete) {
  const card = document.createElement("div");
  card.className = "ext-card";

  const header = document.createElement("div");
  header.className = "ext-card-header";

  const title = document.createElement("span");
  const rawId = ext.id ?? "(unknown extension)";
  title.textContent = rawId.replace(/^Microsoft\.OutlookServices\.OpenTypeExtension\./, "");
  header.appendChild(title);

  // Copy JSON button
  const copyBtn = document.createElement("button");
  copyBtn.className = "ext-card-copy";
  copyBtn.textContent = "Copy";
  copyBtn.addEventListener("click", async (ev) => {
    ev.stopPropagation();
    try {
      await navigator.clipboard.writeText(JSON.stringify(ext, null, 2));
      copyBtn.textContent = "Copied!";
      copyBtn.classList.add("copied");
      setTimeout(() => {
        copyBtn.textContent = "Copy";
        copyBtn.classList.remove("copied");
      }, 1500);
    } catch (_) {
      copyBtn.textContent = "Failed";
      setTimeout(() => { copyBtn.textContent = "Copy"; }, 1500);
    }
  });
  header.appendChild(copyBtn);

  if (onDelete) {
    const delBtn = document.createElement("button");
    delBtn.className = "ext-card-delete";
    delBtn.textContent = "Delete";
    let pendingConfirm = false;

    const resetBtn = () => {
      pendingConfirm = false;
      delBtn.textContent = "Delete";
      delBtn.classList.remove("confirm");
    };

    delBtn.addEventListener("click", async (ev) => {
      ev.stopPropagation();
      if (!pendingConfirm) {
        pendingConfirm = true;
        delBtn.textContent = "Confirm?";
        delBtn.classList.add("confirm");
        const outsideHandler = (ev) => {
          if (!delBtn.contains(ev.target)) {
            resetBtn();
            document.removeEventListener("click", outsideHandler);
          }
        };
        document.addEventListener("click", outsideHandler);
        return;
      }
      delBtn.disabled = true;
      delBtn.textContent = "Deleting…";
      try {
        await onDelete(rawId);
        card.remove();
        const container = document.getElementById("extensions-container");
        if (container.children.length === 0) {
          const empty = document.createElement("p");
          empty.className = "empty";
          empty.textContent = "No extensions found on this message.";
          container.appendChild(empty);
        }
      } catch (e) {
        resetBtn();
        delBtn.disabled = false;
        setStatus(e.message || "Delete failed.", true);
      }
    });

    header.appendChild(delBtn);
  }

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

function appendExtension(ext, onDelete) {
  document.getElementById("extensions-container").appendChild(renderExtension(ext, onDelete));
}

// ── Entry point ───────────────────────────────────────────────────────────────
Office.onReady(async () => {
  const item = Office.context.mailbox.item;
  document.getElementById("ei-from").textContent    = item.from?.emailAddress ?? "—";
  document.getElementById("ei-subject").textContent = item.subject ?? "—";

  document.getElementById("signout-btn").addEventListener("click", (e) => {
    e.preventDefault();
    clearCachedToken();
    location.reload();
  });

  const restId = Office.context.mailbox.convertToRestId(
    item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );

  let token;
  try {
    setStatus("Authenticating…");
    token = await getToken();

    setStatus("Loading extensions…");
    const extensions = await fetchAllExtensions(token, restId);

    if (extensions.length === 0) {
      setStatus("");
      const empty = document.createElement("p");
      empty.className = "empty";
      empty.textContent = "No extensions found on this message.";
      document.getElementById("extensions-container").appendChild(empty);
      return;
    }

    setStatus("");
    extensions.forEach(ext => {
      const onDelete = (extId) => deleteGraphExtension(token, restId, extId);
      appendExtension(ext, onDelete);
    });
  } catch (e) {
    if (e.status === 401) {
      clearCachedToken();
      setStatus("Session expired. Close and reopen the add-in to sign in again.", true);
    } else if (e.status === 405) {
      // Personal MSA accounts don't support the extensions collection endpoint.
      // Fall back to name-based lookup.
      setStatus("");
      document.getElementById("msa-fallback").style.display = "block";

      const input     = document.getElementById("ext-name-input");
      const container = document.getElementById("extensions-container");

      // Restore last-used name; fall back to default
      const savedName = localStorage.getItem(LAST_EXT_KEY);
      input.value = savedName || DEFAULT_EXT;

      const doLookup = async () => {
        const name = input.value.trim();
        if (!name) return;
        setStatus("Looking up…");
        container.innerHTML = "";
        try {
          const ext = await fetchExtensionByName(token, restId, name);
          setStatus("");
          if (!ext) {
            setStatus(`Not found: ${name}`, true);
          } else {
            localStorage.setItem(LAST_EXT_KEY, name);
            const onDelete = async (extId) => {
              await deleteGraphExtension(token, restId, extId);
              setStatus("Extension deleted.");
            };
            appendExtension(ext, onDelete);
          }
        } catch (lookupErr) {
          setStatus(lookupErr.message || "Lookup failed.", true);
        }
      };

      document.getElementById("lookup-btn").addEventListener("click", doLookup);
      if (input.value.trim()) doLookup();
    } else {
      setStatus(e.message || "Failed to load extensions.", true);
    }
  }
});
