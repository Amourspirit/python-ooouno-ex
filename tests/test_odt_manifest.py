import pytest
import xml.dom.minidom


from xml.dom.minicompat import NodeList
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

def test_insert_scrip(basic_odt_manifest_path: Path):
    domtree: xml.dom.minidom.Document = xml.dom.minidom.parse(str(basic_odt_manifest_path))
    assert domtree is not None
    group: xml.dom.minidom.Element = domtree.documentElement
    ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    nl: NodeList = group.getElementsByTagNameNS(ns, "file-entry")
    assert len(nl) == 8

    el_scripts_full: xml.dom.minidom.Element = domtree.createElementNS(ns, "manifest:file-entry")
    el_scripts_full.setAttributeNS(ns, "manifest:full-path", "Scripts/python/script.py")
    el_scripts_full.setAttributeNS(ns, "manifest:media-type", "")
    group.appendChild(el_scripts_full)

    el_scripts_python: xml.dom.minidom.Element = domtree.createElementNS(ns, "manifest:file-entry")
    el_scripts_python.setAttributeNS(ns, "manifest:full-path", "Scripts/python/")
    el_scripts_python.setAttributeNS(ns, "manifest:media-type", "application/binary")
    group.appendChild(el_scripts_python)
    
    el_scripts: xml.dom.minidom.Element = domtree.createElementNS(ns, "manifest:file-entry")
    el_scripts.setAttributeNS(ns, "manifest:full-path", "Scripts/")
    el_scripts.setAttributeNS(ns, "manifest:media-type", "application/binary")
    group.appendChild(el_scripts)

    nl: NodeList = group.getElementsByTagNameNS(ns, "file-entry")
    assert len(nl) == 11
    # x_txt = domtree.toprettyxml()
    return