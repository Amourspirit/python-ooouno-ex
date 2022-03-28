import pytest
import xml.dom.minidom
from pathlib import Path
if __name__ == "__main__":
    pytest.main([__file__])

@pytest.fixture(scope="session")
def basic_odt_manifest_path(fixture_path: Path) -> Path:
    return fixture_path / 'xml' / 'basic_odt_manifest.xml'

def test_load_xml(basic_odt_manifest_path: Path):
    domtree = xml.dom.minidom.parse(str(basic_odt_manifest_path))
    assert domtree is not None
    assert isinstance(domtree, xml.dom.minidom.Document)

def test_get_first(basic_odt_manifest_path: Path):
    domtree: xml.dom.minidom.Document = xml.dom.minidom.parse(str(basic_odt_manifest_path))
    assert domtree is not None