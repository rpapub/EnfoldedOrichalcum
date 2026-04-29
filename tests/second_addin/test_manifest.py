"""Validate docs/addins/second-addin/manifest.xml structure and content."""

import sys
from pathlib import Path
import pytest
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent.parent))
from conftest import ADDIN_MANIFEST_PATHS, ADDIN_BASE_URLS

MANIFEST = ADDIN_MANIFEST_PATHS["second-addin"]
BASE_URL = ADDIN_BASE_URLS["second-addin"]

NS = {
    "o":  "http://schemas.microsoft.com/office/appforoffice/1.1",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
    "vo": "http://schemas.microsoft.com/office/mailappversionoverrides",
}


@pytest.fixture(scope="module")
def tree():
    return etree.parse(str(MANIFEST))


def test_manifest_parses(tree):
    assert tree is not None


def test_required_top_level_fields(tree):
    root = tree.getroot()
    for tag in ("Id", "Version", "ProviderName", "DisplayName", "Description"):
        el = root.find(f"o:{tag}", NS)
        assert el is not None, f"<{tag}> missing"
        assert (el.text or el.get("DefaultValue", "")).strip(), f"<{tag}> is empty"


def test_source_locations_are_https(tree):
    for el in tree.iter():
        for attr in ("DefaultValue", "SourceLocation"):
            val = el.get(attr, "")
            if val.startswith("http"):
                assert val.startswith("https://"), f"Non-HTTPS URL: {val}"


def test_source_location_points_to_base(tree):
    root = tree.getroot()
    src = root.find(".//o:SourceLocation", NS)
    assert src is not None
    assert src.get("DefaultValue", "").startswith(BASE_URL)


def test_icon_url_present(tree):
    root = tree.getroot()
    icon = root.find("o:IconUrl", NS)
    assert icon is not None
    assert icon.get("DefaultValue", "").startswith("https://")


def test_permissions(tree):
    root = tree.getroot()
    perm = root.find("o:Permissions", NS)
    assert perm is not None
    assert perm.text in ("ReadItem", "ReadWriteMailbox", "ReadWriteItem")


def test_version_overrides_present(tree):
    root = tree.getroot()
    vo = root.find("vo:VersionOverrides", NS)
    assert vo is not None, "<VersionOverrides> missing"


def test_bt_images_have_lowercase_size(tree):
    """bt:Image inside <Icon> must use 'size' (lowercase), not 'Size'."""
    for img in tree.iter("{http://schemas.microsoft.com/office/officeappbasictypes/1.0}Image"):
        assert "Size" not in img.attrib, (
            f"bt:Image uses 'Size' (capital) — must be lowercase 'size'. "
            f"Attributes: {dict(img.attrib)}"
        )


def test_bt_images_icon_have_size(tree):
    """bt:Image elements inside <Icon> controls must have a 'size' attribute."""
    vo = tree.find("vo:VersionOverrides", NS)
    assert vo is not None
    for icon in vo.iter("{http://schemas.microsoft.com/office/mailappversionoverrides}Icon"):
        for img in icon.iter("{http://schemas.microsoft.com/office/officeappbasictypes/1.0}Image"):
            assert "size" in img.attrib, f"bt:Image inside <Icon> missing 'size': {dict(img.attrib)}"
            assert img.get("size") in ("16", "32", "80"), f"Unexpected icon size: {img.get('size')}"
