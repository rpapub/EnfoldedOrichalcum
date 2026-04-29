"""Playwright tests for embed-data browser preview (no Office.js needed)."""
import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page, expect

sys.path.insert(0, str(Path(__file__).parent.parent))
from conftest import ADDIN_BASE_URLS

BASE = ADDIN_BASE_URLS["embed-data"]


def test_preview_form_renders(page: Page):
    page.goto(f"{BASE}/", wait_until="domcontentloaded")
    expect(page.locator("#f-reference")).to_be_visible()
    expect(page.locator("#f-country")).to_be_visible()


def test_country_dropdown_populated(page: Page):
    page.goto(f"{BASE}/", wait_until="domcontentloaded")
    options = page.locator("#f-country option")
    expect(options).to_have_count(32)  # "— select —" + 31 countries


def test_save_preview_shows_json(page: Page):
    page.goto(f"{BASE}/", wait_until="domcontentloaded")
    page.fill("#f-reference", "TEST-001")
    page.fill("#f-category", "Support")
    page.click("#save-btn")
    output = page.locator("#preview-output")
    expect(output).to_be_visible()
    expect(output).to_contain_text("TEST-001")
    expect(output).to_contain_text("Support")
