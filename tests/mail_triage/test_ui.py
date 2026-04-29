"""Playwright UI tests against the live GitHub Pages testbed for mail-triage."""

import json
import re
import sys
from pathlib import Path

import pytest
from playwright.sync_api import Page, expect

sys.path.insert(0, str(Path(__file__).parent.parent))
from conftest import ADDIN_BASE_URLS

BASE = ADDIN_BASE_URLS["mail-triage"]


@pytest.fixture(scope="module")
def testbed(browser):
    page = browser.new_page()
    page.goto(f"{BASE}/", wait_until="networkidle")
    yield page
    page.close()


def test_testbed_title(testbed):
    expect(testbed).to_have_title(re.compile("Mail Triage", re.IGNORECASE))


def test_testbed_form_fields_present(testbed):
    for field_id in ("f-case-id", "f-case-status", "f-state-version",
                     "f-triage-result", "f-triage-confidence", "f-triage-summary"):
        expect(testbed.locator(f"#{field_id}")).to_be_visible()


def test_case_status_dropdown_populated(testbed):
    options = testbed.locator("#f-case-status option")
    assert options.count() >= 5


def test_case_status_includes_triaged(testbed):
    option = testbed.locator("#f-case-status option[value='triaged']")
    expect(option).to_have_count(1)


def test_attachment_profile_field_present(testbed):
    # attachmentProfile has x-default:true in schema so it renders automatically
    field = testbed.locator('[data-ev-key="attachmentProfile"]')
    expect(field).to_be_visible()


def test_attachment_profile_includes_pdf_and_excel(testbed):
    option = testbed.locator('[data-ev-key="attachmentProfile"] option[value="pdf_and_excel"]')
    expect(option).to_have_count(1)


def test_preview_output_shows_nested_json(testbed):
    testbed.locator("#save-btn").click()
    output = testbed.locator("#preview-output")
    expect(output).to_be_visible()
    raw = output.text_content()
    data = json.loads(raw)
    assert "caseId" in data
    assert "caseStatus" in data
    assert "stateVersion" in data
    assert "triage" in data
    assert "result" in data["triage"]
    assert "confidence" in data["triage"]


def test_preview_triage_defaults(testbed):
    testbed.locator("#save-btn").click()
    raw = testbed.locator("#preview-output").text_content()
    data = json.loads(raw)
    assert data["triage"]["result"] == "no_action"
    assert data["triage"]["confidence"] == "low"
    assert "evidence" not in data  # no evidence for an untriaged preview


def test_manifest_xml_reachable(page: Page):
    response = page.goto(f"{BASE}/manifest.xml")
    assert response.status == 200
    assert "xml" in response.headers.get("content-type", "")


def test_taskpane_html_loads(page: Page):
    page.goto(f"{BASE}/taskpane.html", wait_until="domcontentloaded")
    expect(page.locator("#save-btn")).to_be_visible()


def test_taskpane_has_triage_fields(page: Page):
    page.goto(f"{BASE}/taskpane.html", wait_until="domcontentloaded")
    for field_id in ("f-case-id", "f-case-status", "f-triage-result",
                     "f-triage-confidence", "f-triage-summary"):
        expect(page.locator(f"#{field_id}")).to_be_attached()
