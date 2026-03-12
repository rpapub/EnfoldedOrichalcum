# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install Python deps + Playwright browser
just install          # uv sync --group test && playwright install chromium

# Validate the Office manifest XML
just validate         # npx office-addin-manifest validate docs/manifest.xml

# Run all tests
just test             # uv run pytest

# Run only manifest tests (no network)
just test-manifest    # uv run pytest tests/test_manifest.py -v

# Run only UI tests (hits live GitHub Pages)
just test-ui          # uv run pytest tests/test_ui.py -v

# Run a single test
uv run pytest tests/test_manifest.py::test_bt_images_have_lowercase_size -v

# Interactive email triage CLI (opens Firefox in WSL for auth)
just triage

# Read back extension data for a message ID
just check <message-id>
```

## Architecture

This is a static GitHub-Pages-hosted Microsoft Outlook task pane add-in. Everything is served from `docs/` on the `development` branch — no build step, no bundler.

### Auth flow (critical, non-obvious)

Outlook task panes run in Edge WebView2 with Tracking Prevention enabled. This **blocks MSAL's CDN** (`alcdn.msauth.net`). There is no MSAL in the task pane at all.

Authentication is delegated entirely to `docs/auth-dialog.html`, which runs in a real browser window opened via `Office.context.ui.displayDialogAsync`. It implements a **pure PKCE OAuth2** flow using only the Web Crypto API and `fetch` — no external libraries. The dialog sends the token back to the task pane via `Office.context.ui.messageParent`.

The resulting access token is cached in `localStorage` (`triage_token_cache`) with an expiry timestamp. The dialog only opens when the cache is empty or the token expires within 60 seconds.

### Form state persistence

Before opening the auth dialog, the task pane serializes form fields to `localStorage` (`triage_pending_form`). On `Office.onReady`, it restores from that key — this recovers user input if the task pane iframe reloads during the PKCE redirect cycle. The key is cleared after a successful Graph write.

### Graph open extension

Triage data is written as `microsoft.graph.openTypeExtension` named `net.cprima.rpapub.EnfoldedOrichalcum` on the message resource. Write is a POST; if the extension already exists (409), it falls back to PATCH. Read-back uses `$filter=extensions/any(f:f/id eq 'net.cprima.rpapub.EnfoldedOrichalcum')` with `$expand` — this is also how the Python CLI queries tagged messages, because direct `/messages/{id}/extensions/...` URLs fail for EWS IDs that contain `/` characters.

### Message ID formats

`Office.context.mailbox.item.itemId` returns an EWS-format ID (contains literal `/`). Before passing to Graph, it must be converted:

```javascript
const restId = Office.context.mailbox.convertToRestId(
  Office.context.mailbox.item.itemId,
  Office.MailboxEnums.RestVersion.v2_0
);
```

`translateExchangeIds` (Graph endpoint) returns 403 for personal MSA accounts — do not use it.

### Python CLI (`scripts/triage.py`)

Uses MSAL Python (`acquire_token_interactive`) with `http://localhost` redirect URI (registered as Mobile/desktop platform in Azure AD). In WSL, the `BROWSER` env var must be set to the Windows Firefox path (encoded in the justfile). Uses `tty_input()` (reads `/dev/tty` directly) instead of `input()` because `just` closes stdin.

### Tests

- `tests/test_manifest.py` — parses `docs/manifest.xml` with lxml, validates structure, checks `bt:Image` uses lowercase `size` attribute (a real validator gotcha), verifies `EXTENSION_NAME` and client ID in `taskpane.js`. **No network required.**
- `tests/test_ui.py` — Playwright against the live `https://rpapub.github.io/EnfoldedOrichalcum` URL. Requires GitHub Pages to be deployed and up to date.

### Manifest gotchas

- `bt:Image` inside `<Icon>` requires `size` (lowercase), not `Size` — the Office validator rejects capital-S.
- `xmlns:bt` must be redeclared on `<VersionOverrides>` even though it is on the root element.
- `<AppDomains>` must include `login.microsoftonline.com` and `graph.microsoft.com`.
- After any manifest change, the add-in must be removed and re-added in Outlook — a task pane refresh is not sufficient.

## Key constants

| Constant | Location | Value |
|---|---|---|
| `CLIENT_ID` | `taskpane.js`, `auth-dialog.html`, `scripts/triage.py` | `f28629c6-2f87-4afc-a6ff-1cbbd50166af` |
| `EXTENSION_NAME` | same | `net.cprima.rpapub.EnfoldedOrichalcum` |
| `REDIRECT_URI` | `auth-dialog.html` | `https://rpapub.github.io/EnfoldedOrichalcum/auth-dialog.html` |

## Deployment

Push to `development` branch → GitHub Pages redeploys automatically (~1–2 min). The UI tests hit the live deployed URL, so they may fail immediately after a push until Pages finishes rebuilding.
