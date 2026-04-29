/* Shared Graph API utilities — loaded by each add-in's taskpane.html */

const TOKEN_KEY = "triage_token_cache";

function getCachedToken() {
  try {
    const cached = JSON.parse(localStorage.getItem(TOKEN_KEY) || "null");
    if (cached && cached.expiresAt > Date.now() + 60_000) return cached.accessToken;
  } catch (_) {}
  return null;
}

function cacheToken(accessToken, expiresIn) {
  localStorage.setItem(TOKEN_KEY, JSON.stringify({
    accessToken,
    expiresAt: Date.now() + expiresIn * 1000
  }));
}

function clearCachedToken() {
  localStorage.removeItem(TOKEN_KEY);
}

function graphError(status, text) {
  const err = new Error(text || `Graph error ${status}`);
  err.status = status;
  return err;
}

async function readGraphExtension(token, restMessageId, extensionName) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions/${encodeURIComponent(extensionName)}`;
  console.log("[graph] GET", url);
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  console.log("[graph] GET status", res.status);
  if (res.status === 404) return null;
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    console.error("[graph] GET error body", detail);
    throw graphError(res.status, `Graph ${res.status}`);
  }
  const json = await res.json();
  console.log("[graph] GET response", json);
  return json;
}

async function deleteGraphExtension(token, restMessageId, extensionId) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions/${encodeURIComponent(extensionId)}`;
  const res = await fetch(url, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 204 || res.ok) return;
  const text = await res.text().catch(() => "");
  throw graphError(res.status, text);
}

async function writeGraphExtension(token, restMessageId, extensionName, data) {
  const baseUrl = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions`;
  const patchUrl = `${baseUrl}/${extensionName}`;
  const payload = {
    "@odata.type": "microsoft.graph.openTypeExtension",
    id: extensionName,
    extensionName,
    ...data
  };
  const body = JSON.stringify(payload);
  const headers = {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json"
  };

  console.log("[graph] POST", baseUrl);
  console.log("[graph] body", payload);

  let res = await fetch(baseUrl, { method: "POST", headers, body });
  console.log("[graph] POST status", res.status);

  // 409 = already exists; 500 = some Exchange backends return this instead of 409
  if (res.status === 409 || res.status === 500) {
    console.log(`[graph] ${res.status} → PATCH`, patchUrl);
    res = await fetch(patchUrl, { method: "PATCH", headers, body });
    console.log("[graph] PATCH status", res.status);
    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      console.error("[graph] PATCH error body", detail);
      throw graphError(res.status, `Extension update failed: ${res.status}`);
    }
    return;
  }

  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    console.error("[graph] POST error body", detail);
    throw graphError(res.status, detail);
  }
}

/* istanbul ignore next -- Node.js/test export only, no effect in browsers */
if (typeof module !== "undefined") {
  module.exports = {
    getCachedToken, cacheToken, clearCachedToken,
    graphError, readGraphExtension, writeGraphExtension, deleteGraphExtension
  };
}
