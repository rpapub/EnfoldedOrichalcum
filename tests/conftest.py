from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent

ADDIN_MANIFEST_PATHS = {
    "embed-data":          REPO_ROOT / "docs/addins/embed-data/manifest.xml",
    "second-addin":        REPO_ROOT / "docs/addins/second-addin/manifest.xml",
    "extension-inspector": REPO_ROOT / "docs/addins/extension-inspector/manifest.xml",
}

ADDIN_BASE_URLS = {
    "embed-data":          "https://rpapub.github.io/EnfoldedOrichalcum/addins/embed-data",
    "second-addin":        "https://rpapub.github.io/EnfoldedOrichalcum/addins/second-addin",
    "extension-inspector": "https://rpapub.github.io/EnfoldedOrichalcum/addins/extension-inspector",
}
