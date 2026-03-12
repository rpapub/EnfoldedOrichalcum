"""Playwright UI tests against the live GitHub Pages testbed."""

import json
import re
from datetime import date

import pytest
from playwright.sync_api import Page, expect

BASE = "https://rpapub.github.io/EnfoldedOrichalcum"
CURRENT_YEAR = str(date.today().year)
CURRENT_QUARTER = f"Q{(date.today().month - 1) // 3 + 1}"


# ── index.html testbed ────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def testbed(browser):
    page = browser.new_page()
    page.goto(f"{BASE}/", wait_until="networkidle")
    yield page
    page.close()


def test_testbed_title(testbed):
    expect(testbed).to_have_title(re.compile("Email Triage", re.IGNORECASE))


def test_testbed_form_fields_present(testbed):
    for field_id in ("from", "subject", "f-reference", "f-category",
                     "f-priority", "f-owner", "f-notes"):
        expect(testbed.locator(f"#{field_id}")).to_be_visible()


def test_testbed_country_dropdown_populated(testbed):
    options = testbed.locator("#f-country option")
    assert options.count() >= 30


def test_testbed_country_includes_usa(testbed):
    option = testbed.locator("#f-country option[value='USA']")
    expect(option).to_have_count(1)


def test_testbed_year_dropdown_has_current_year(testbed):
    first = testbed.locator("#f-year option").first
    expect(first).to_have_attribute("value", CURRENT_YEAR)


def test_testbed_year_dropdown_has_six_years(testbed):
    assert testbed.locator("#f-year option").count() == 6


def test_testbed_quarter_default(testbed):
    value = testbed.locator("#f-quarter").input_value()
    assert value == CURRENT_QUARTER


def test_testbed_preview_output_shows_json(testbed):
    testbed.locator("#save-btn").click()
    output = testbed.locator("#preview-output")
    expect(output).to_be_visible()
    raw = output.text_content()
    data = json.loads(raw)
    assert "reference" in data
    assert "country" in data
    assert "year" in data
    assert "quarter" in data
    assert data["year"] == CURRENT_YEAR
    assert data["quarter"] == CURRENT_QUARTER


def test_testbed_preview_from_field(testbed):
    testbed.locator("#save-btn").click()
    raw = testbed.locator("#preview-output").text_content()
    data = json.loads(raw)
    assert data["from"] == "sender@example.com"


# ── manifest.xml reachable ────────────────────────────────────────────────────

def test_manifest_xml_reachable(page: Page):
    response = page.goto(f"{BASE}/manifest.xml")
    assert response.status == 200
    assert "xml" in response.headers.get("content-type", "")


# ── taskpane.html loads ───────────────────────────────────────────────────────

def test_taskpane_html_loads(page: Page):
    page.goto(f"{BASE}/taskpane.html", wait_until="domcontentloaded")
    expect(page.locator("#save-btn")).to_be_visible()


def test_taskpane_has_triage_fields(page: Page):
    page.goto(f"{BASE}/taskpane.html", wait_until="domcontentloaded")
    for field_id in ("f-reference", "f-category", "f-priority", "f-owner", "f-notes",
                     "f-country", "f-year", "f-quarter"):
        expect(page.locator(f"#{field_id}")).to_be_attached()
