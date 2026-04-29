# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install Python deps + Playwright browser
just install          # uv sync --group test && playwright install chromium

# Validate a specific add-in manifest
just validate                  # validates embed-data (default)
just validate extension-inspector

# Validate all manifests
just validate-all

# Run all tests
just test             # uv run pytest

# Run only manifest tests (no network)
just test-manifest                  # all add-ins
just test-manifest embed_data       # embed-data only

# Run only UI tests (hits live GitHub Pages)
just test-ui                        # all add-ins
just test-ui embed_data             # embed-data only

# Run a single test
uv run pytest tests/embed_data/test_manifest.py::test_bt_images_have_lowercase_size -v

# Serve docs/ locally for development
just serve             # python -m http.server 3000 --directory docs

# Lint JavaScript
just lint              # npx eslint docs/addins/**/*.js docs/shared/graph-utils.js

# Run JavaScript unit tests (Vitest)
just test-unit         # npx vitest run --reporter=verbose

# Run browser preview Playwright tests (no Outlook context needed)
just test-ui-preview   # tests/embed_data/test_ui_preview.py + tests/extension_inspector/test_ui_mocked.py

# Interactive email triage CLI (opens Firefox in WSL for auth)
just triage

# Read back extension data for a message ID
just check <message-id>
```

## Monorepo layout

This repo hosts N Outlook Add-ins, all served as static files from `docs/` on the `development` branch — no build step, no bundler. GitHub Pages serves the whole `docs/` tree as-is.

```
docs/
├── index.html                        # landing page linking to each add-in
├── shared/
│   ├── auth-dialog.html              # shared PKCE OAuth2 dialog
│   ├── solarized.css                 # shared Solarized Light CSS variables
│   └── graph-utils.js                # shared: getCachedToken, cacheToken, clearCachedToken, graphError, writeGraphExtension
└── addins/
    ├── embed-data/                   # RPA Embed Data add-in
    │   ├── index.html                # testbed + install guide
    │   ├── taskpane.html             # Outlook task pane UI
    │   ├── taskpane.js               # add-in logic (uses shared/graph-utils.js globals)
    │   ├── manifest.xml              # all URLs point to /addins/embed-data/
    │   └── assets/icon-{16,32,80}.png
    └── extension-inspector/          # Extension Inspector add-in
        ├── index.html
        ├── taskpane.html
        ├── taskpane.js
        ├── manifest.xml
        └── assets/icon-{16,32,80}.png

scripts/
└── embed_data_triage.py              # Python CLI for embed-data (MSAL, Graph API)

tests/
├── conftest.py                       # ADDIN_BASE_URLS + ADDIN_MANIFEST_PATHS registry
├── unit/
│   └── graph-utils.test.js           # Vitest unit tests for graph-utils.js
├── embed_data/
│   ├── test_manifest.py              # XML validation (no network)
│   ├── test_ui.py                    # Playwright (live GitHub Pages)
│   └── test_ui_preview.py            # Playwright (browser preview, no Office.js)
└── extension_inspector/
    ├── test_ui_mocked.py             # Playwright with mocked Graph (xfail until Office.js stub)
```

### Adding a new add-in

1. Create `docs/addins/<slug>/` with `index.html`, `taskpane.html`, `taskpane.js`, `manifest.xml`, `assets/`
2. Generate a new UUID for `<Id>` in `manifest.xml` (`python -c "import uuid; print(uuid.uuid4())"`)
3. Set all manifest URLs to `/addins/<slug>/` paths
4. If using the same Azure AD app registration: point `DIALOG_URL` in `taskpane.js` to `shared/auth-dialog.html`. If a new registration is needed: add a separate `auth-dialog.html` in the addin directory and register its URL as a redirect URI in Azure AD
5. Add the slug to `ADDIN_BASE_URLS` and `ADDIN_MANIFEST_PATHS` in `tests/conftest.py`
6. Add manifest validation step to `pages.yml` (optional but recommended)
7. Add the addin card to `docs/index.html`
8. Create `tests/<slug>/test_manifest.py` and `tests/<slug>/test_ui.py`

### Migration note

The old root-level URLs (`/taskpane.html`, `/manifest.xml`, `/auth-dialog.html`) no longer exist. Anyone who installed the old manifest must **uninstall and reinstall** using the new URL: `https://rpapub.github.io/EnfoldedOrichalcum/addins/embed-data/manifest.xml`

## Architecture

### Auth flow (critical, non-obvious)

Outlook task panes run in Edge WebView2 with Tracking Prevention enabled. This **blocks MSAL's CDN** (`alcdn.msauth.net`). There is no MSAL in the task pane at all.

Authentication is delegated entirely to `docs/shared/auth-dialog.html`, which runs in a real browser window opened via `Office.context.ui.displayDialogAsync`. It implements a **pure PKCE OAuth2** flow using only the Web Crypto API and `fetch` — no external libraries. The dialog sends the token back to the task pane via `Office.context.ui.messageParent`.

The shared `auth-dialog.html` works for all add-ins that share the same Azure AD app registration (same `CLIENT_ID`). If a second add-in needs a different Azure AD registration, it gets its own `auth-dialog.html` in its addin directory.

The resulting access token is cached in `localStorage` (`triage_token_cache`) with an expiry timestamp. The dialog only opens when the cache is empty or the token expires within 60 seconds.

### Form state persistence

Before opening the auth dialog, the task pane serializes form fields to `localStorage` (`triage_pending_form`). On `Office.onReady`, it restores from that key — this recovers user input if the task pane iframe reloads during the PKCE redirect cycle. The key is cleared after a successful Graph write.

### Shared JS utilities (`docs/shared/graph-utils.js`)

Loaded by each add-in's `taskpane.html` before `taskpane.js`. Exposes plain globals:
- `getCachedToken()` — reads `localStorage` token cache
- `cacheToken(accessToken, expiresIn)` — writes token cache
- `clearCachedToken()` — removes the cached token (used by sign-out and 401 recovery)
- `graphError(status, text)` — creates an `Error` with `.status` attached, for status-code branching in catch blocks
- `readGraphExtension(token, restMessageId, extensionName)` — GET single extension by name; returns null on 404
- `deleteGraphExtension(token, restMessageId, extensionId)` — DELETE extension; resolves on 204
- `writeGraphExtension(token, restMessageId, extensionName, data)` — POST/PATCH Graph open extension

`extensionName` is passed by the caller so each add-in can use its own extension name.

The file also exports all functions via `module.exports` when running in Node.js (for Vitest tests), with no effect in browsers.

### Graph open extension

Triage data is written as `microsoft.graph.openTypeExtension` on the message resource. Write is a POST; if the extension already exists (409), it falls back to PATCH. Read-back uses `$filter=extensions/any(f:f/id eq '...')` with `$expand` — this is also how the Python CLI queries tagged messages, because direct `/messages/{id}/extensions/...` URLs fail for EWS IDs that contain `/` characters.

### Message ID formats

`Office.context.mailbox.item.itemId` returns an EWS-format ID (contains literal `/`). Before passing to Graph, it must be converted:

```javascript
const restId = Office.context.mailbox.convertToRestId(
  Office.context.mailbox.item.itemId,
  Office.MailboxEnums.RestVersion.v2_0
);
```

`translateExchangeIds` (Graph endpoint) returns 403 for personal MSA accounts — do not use it.

### Python CLI (`scripts/embed_data_triage.py`)

Uses MSAL Python (`acquire_token_interactive`) with `http://localhost` redirect URI (registered as Mobile/desktop platform in Azure AD). In WSL, the `BROWSER` env var must be set to the Windows Firefox path (encoded in the justfile). Uses `tty_input()` (reads `/dev/tty` directly) instead of `input()` because `just` closes stdin.

### Tests

- `tests/embed_data/test_manifest.py` — parses `docs/addins/embed-data/manifest.xml` with lxml, validates structure, checks `bt:Image` uses lowercase `size` attribute (a real validator gotcha), verifies `EXTENSION_NAME` and client ID in `taskpane.js`. **No network required.**
- `tests/embed_data/test_ui.py` — Playwright against the live `https://rpapub.github.io/EnfoldedOrichalcum/addins/embed-data` URL. Requires GitHub Pages to be deployed and up to date.
- `tests/embed_data/test_ui_preview.py` — Playwright against the embed-data index.html browser preview (no Office.js). Covers form rendering, dropdown population, and save-button preview output.
- `tests/extension_inspector/test_ui_mocked.py` — Playwright with mocked Graph API. Marked `xfail` until an Office.js stub is injected via `page.add_init_script`.
- `tests/unit/graph-utils.test.js` — Vitest unit tests for `docs/shared/graph-utils.js`. Run with `just test-unit`.
- `tests/conftest.py` — central registry of addin slugs, manifest paths, and live base URLs.

### Manifest gotchas

- `bt:Image` inside `<Icon>` requires `size` (lowercase), not `Size` — the Office validator rejects capital-S.
- `xmlns:bt` must be redeclared on `<VersionOverrides>` even though it is on the root element.
- `<AppDomains>` must include `login.microsoftonline.com` and `graph.microsoft.com`. A single `https://rpapub.github.io` entry covers all subdirectory paths.
- After any manifest change, the add-in must be removed and re-added in Outlook — a task pane refresh is not sufficient.

## Key constants

| Constant | Add-in | Location | Value |
|---|---|---|---|
| `CLIENT_ID` | embed-data | `taskpane.js`, `shared/auth-dialog.html`, `scripts/embed_data_triage.py` | `f28629c6-2f87-4afc-a6ff-1cbbd50166af` |
| `EXTENSION_NAME` | embed-data | `taskpane.js` (passed to graph-utils.js) | `net.cprima.rpapub.EnfoldedOrichalcum` |
| `REDIRECT_URI` | shared | `docs/shared/auth-dialog.html` | `https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html` |
| `DIALOG_URL` | embed-data | `docs/addins/embed-data/taskpane.js` | same as `REDIRECT_URI` above |

## Deployment

Push to `development` branch → GitHub Pages redeploys automatically (~1–2 min). The UI tests hit the live deployed URL, so they may fail immediately after a push until Pages finishes rebuilding.

Manifest can also be triggered manually: `gh workflow run pages.yml`
