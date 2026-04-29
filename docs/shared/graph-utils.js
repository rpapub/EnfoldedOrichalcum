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

async function readGraphExtension(token, restMessageId, extensionName) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions/${encodeURIComponent(extensionName)}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Graph ${res.status}`);
  return res.json();
}

async function deleteGraphExtension(token, restMessageId, extensionId) {
  const url = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions/${encodeURIComponent(extensionId)}`;
  const res = await fetch(url, { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 204 || res.ok) return;
  const text = await res.text().catch(() => "");
  const err = new Error(text || `Graph error ${res.status}`);
  err.status = res.status;
  throw err;
}

async function writeGraphExtension(token, restMessageId, extensionName, data) {
  const baseUrl = `https://graph.microsoft.com/v1.0/me/messages/${restMessageId}/extensions`;
  const body = JSON.stringify({
    "@odata.type": "microsoft.graph.openTypeExtension",
    extensionName,
    ...data
  });
  const headers = {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json"
  };

  let res = await fetch(baseUrl, { method: "POST", headers, body });

  if (res.status === 409) {
    res = await fetch(`${baseUrl}/${extensionName}`, { method: "PATCH", headers, body });
    if (res.status !== 204) throw new Error(`Extension update failed: ${res.status}`);
    return;
  }

  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new Error(`Graph API ${res.status}: ${detail}`);
  }
}
