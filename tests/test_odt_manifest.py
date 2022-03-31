import pytest
import xml.dom.minidom
from xml.dom.minicompat import NodeList
from pathlib import Path
from typing import List

if __name__ == "__main__":
    pytest.main([__file__])


@pytest.fixture(scope="session")
def basic_odt_manifest_path(fixture_path: Path) -> Path:
    return fixture_path / "xml" / "basic_odt_manifest.xml"

@pytest.fixture(scope="session")
def script_odt_manifest_path(fixture_path: Path) -> Path:
    return fixture_path / "xml" / "script_odt_manifest.xml"

def test_load_basic_xml(script_odt_manifest_path: Path):
    domtree = xml.dom.minidom.parse(str(script_odt_manifest_path))
    assert domtree is not None
    assert isinstance(domtree, xml.dom.minidom.Document)

def test_load_script_xml(basic_odt_manifest_path: Path):
    domtree = xml.dom.minidom.parse(str(basic_odt_manifest_path))
    assert domtree is not None
    assert isinstance(domtree, xml.dom.minidom.Document)

def test_insert_script_basic(basic_odt_manifest_path: Path):
    domtree: xml.dom.minidom.Document = xml.dom.minidom.parse(
        str(basic_odt_manifest_path)
    )
    assert domtree is not None
    group: xml.dom.minidom.Element = domtree.documentElement
    ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    nl: NodeList = group.getElementsByTagNameNS(ns, "file-entry")
    assert len(nl) == 8

    el_scripts_full: xml.dom.minidom.Element = domtree.createElementNS(
        ns, "manifest:file-entry"
    )
    el_scripts_full.setAttributeNS(ns, "manifest:full-path", "Scripts/python/script.py")
    el_scripts_full.setAttributeNS(ns, "manifest:media-type", "")
    group.appendChild(el_scripts_full)

    el_scripts_python: xml.dom.minidom.Element = domtree.createElementNS(
        ns, "manifest:file-entry"
    )
    el_scripts_python.setAttributeNS(ns, "manifest:full-path", "Scripts/python/")
    el_scripts_python.setAttributeNS(ns, "manifest:media-type", "application/binary")
    group.appendChild(el_scripts_python)

    el_scripts: xml.dom.minidom.Element = domtree.createElementNS(
        ns, "manifest:file-entry"
    )
    el_scripts.setAttributeNS(ns, "manifest:full-path", "Scripts/")
    el_scripts.setAttributeNS(ns, "manifest:media-type", "application/binary")
    group.appendChild(el_scripts)

    nl: NodeList = group.getElementsByTagNameNS(ns, "file-entry")
    assert len(nl) == 11
    # x_txt = domtree.toprettyxml()
    return

def test_match_script(script_odt_manifest_path):
    ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    def contains(lst: List[xml.dom.minidom.Element], full_path:str, media_type: str = "application/binary") -> bool:
        result = False
        for el in lst:
            fp: xml.dom.minidom.Attr = el.getAttributeNodeNS(ns, "full-path")
            if not fp:
                continue
            mt: xml.dom.minidom.Attr = el.getAttributeNodeNS(ns, "media-type")
            if mt is None:
                continue
            if fp.value == full_path and mt.value == media_type:
                result = True
                break
            
        return result

    domtree: xml.dom.minidom.Document = xml.dom.minidom.parse(
        str(script_odt_manifest_path)
    )
    assert domtree is not None
    group: xml.dom.minidom.Element = domtree.documentElement
    nl: NodeList = group.getElementsByTagNameNS(ns, "file-entry")
    assert len(nl) == 11
    assert contains(nl, "Scripts/") is True
    assert contains(nl, "Scripts/python/") is True
    assert contains(nl, "Scripts/python/script.py", "") == True
    assert contains(nl, "Scripts/Python/") is False
    return
    