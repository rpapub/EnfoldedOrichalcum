"""Playwright UI tests against the live GitHub Pages testbed for extension-inspector."""

import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page, expect

sys.path.insert(0, str(Path(__file__).parent.parent))
from conftest import ADDIN_BASE_URLS

BASE = ADDIN_BASE_URLS["extension-inspector"]


def test_manifest_xml_reachable(page: Page):
    response = page.goto(f"{BASE}/manifest.xml")
    assert response.status == 200
    assert "xml" in response.headers.get("content-type", "")


def test_taskpane_html_loads(page: Page):
    page.goto(f"{BASE}/taskpane.html", wait_until="domcontentloaded")
    expect(page.locator("#extensions-container")).to_be_attached()


def test_index_html_loads(page: Page):
    page.goto(f"{BASE}/", wait_until="domcontentloaded")
    expect(page.locator("h1")).to_contain_text("Extension Inspector")
