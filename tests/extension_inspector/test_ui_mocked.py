"""Playwright tests for extension-inspector with mocked Graph responses."""
import json
import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page, expect

sys.path.insert(0, str(Path(__file__).parent.parent))
from conftest import ADDIN_BASE_URLS

BASE = ADDIN_BASE_URLS["extension-inspector"]

FIXTURE_EXT = {
    "id": "net.cprima.rpapub.EnfoldedOrichalcum",
    "@odata.type": "microsoft.graph.openTypeExtension",
    "reference": "TICKET-42",
    "priority": "High",
}


# NOTE: These tests require a local Office.js stub injected before the page
# scripts run. Until that stub is added, mark as xfail.
@pytest.mark.xfail(reason="Requires Office.js stub — future work")
def test_extension_card_renders(page: Page):
    page.route("**/graph.microsoft.com/**", lambda r: r.fulfill(
        status=200,
        content_type="application/json",
        body=json.dumps({"value": [FIXTURE_EXT]}),
    ))
    page.goto(f"{BASE}/taskpane.html", wait_until="domcontentloaded")
    expect(page.locator(".ext-card")).to_be_visible()
    expect(page.locator(".ext-card-header span")).to_contain_text("EnfoldedOrichalcum")
