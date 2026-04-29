install:
    uv sync --group test
    uv run playwright install chromium

validate addin="embed-data":
    npx office-addin-manifest validate docs/addins/{{addin}}/manifest.xml

validate-all:
    npx office-addin-manifest validate docs/addins/embed-data/manifest.xml
    npx office-addin-manifest validate docs/addins/second-addin/manifest.xml

test:
    uv run pytest

test-manifest addin="":
    #!/usr/bin/env bash
    if [ -z "{{addin}}" ]; then
        uv run pytest tests/ -k "manifest" -v
    else
        uv run pytest tests/{{addin}}/test_manifest.py -v
    fi

test-ui addin="":
    #!/usr/bin/env bash
    if [ -z "{{addin}}" ]; then
        uv run pytest tests/ -k "ui" -v
    else
        uv run pytest tests/{{addin}}/test_ui.py -v
    fi

triage:
    #!/usr/bin/env bash
    exec env BROWSER="/mnt/c/Program Files/Mozilla Firefox/firefox.exe" uv run scripts/embed_data_triage.py

check mid:
    #!/usr/bin/env bash
    exec env BROWSER="/mnt/c/Program Files/Mozilla Firefox/firefox.exe" uv run scripts/embed_data_triage.py --check --message-id "{{mid}}"
