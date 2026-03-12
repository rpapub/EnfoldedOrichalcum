validate:
    npx office-addin-manifest validate docs/manifest.xml

triage:
    #!/usr/bin/env bash
    exec env BROWSER="/mnt/c/Program Files/Mozilla Firefox/firefox.exe" uv run scripts/triage.py

install:
    uv sync --group test
    uv run playwright install chromium

test:
    uv run pytest

test-manifest:
    uv run pytest tests/test_manifest.py -v

test-ui:
    uv run pytest tests/test_ui.py -v
