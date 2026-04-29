from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent

ADDIN_MANIFEST_PATHS = {
    "embed-data":          REPO_ROOT / "docs/addins/embed-data/manifest.xml",
    "extension-inspector": REPO_ROOT / "docs/addins/extension-inspector/manifest.xml",
    "mail-triage":         REPO_ROOT / "docs/addins/mail-triage/manifest.xml",
}

ADDIN_BASE_URLS = {
    "embed-data":          "https://rpapub.github.io/EnfoldedOrichalcum/addins/embed-data",
    "extension-inspector": "https://rpapub.github.io/EnfoldedOrichalcum/addins/extension-inspector",
    "mail-triage":         "https://rpapub.github.io/EnfoldedOrichalcum/addins/mail-triage",
}
