# coding: utf-8
from pathlib import Path
from typing import Optional
import xml.dom.minidom
from config import AppConfig
from ..utils import util


class ManifestScript:
    def __init__(
        self, manifest_path: Path, script_name: str, config: Optional[AppConfig] = None
    ) -> None:
        """
        Constructor

        Args:
            manifest_path (Path): Path to manifest.xml file
            script_name (str): Name of script that is to be embed in document.
            config (Optional[AppConfig], optional): App config. Defaults to None.
        """
        self._path = manifest_path
        self._script_name = script_name
        if config is None:
            self._config = util.get_app_cfg()
        else:
            self._config = config

    def write(self) -> None:
        """
        Adds Script elements to manifest.xml file.
        """
        domtree: xml.dom.minidom.Document = xml.dom.minidom.parse(str(self._path))
        group: xml.dom.minidom.Element = domtree.documentElement
        # ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
        ns = self._config.xml_manifest_namesapce

        el_scripts_full: xml.dom.minidom.Element = domtree.createElementNS(
            ns, "manifest:file-entry"
        )
        el_scripts_full.setAttributeNS(
            ns, "manifest:full-path", f"Scripts/python/{self._script_name}"
        )
        el_scripts_full.setAttributeNS(ns, "manifest:media-type", "")
        group.appendChild(el_scripts_full)

        el_scripts_python: xml.dom.minidom.Element = domtree.createElementNS(
            ns, "manifest:file-entry"
        )
        el_scripts_python.setAttributeNS(ns, "manifest:full-path", "Scripts/python/")
        el_scripts_python.setAttributeNS(
            ns, "manifest:media-type", "application/binary"
        )
        group.appendChild(el_scripts_python)

        el_scripts: xml.dom.minidom.Element = domtree.createElementNS(
            ns, "manifest:file-entry"
        )
        el_scripts.setAttributeNS(ns, "manifest:full-path", "Scripts/")
        el_scripts.setAttributeNS(ns, "manifest:media-type", "application/binary")
        group.appendChild(el_scripts)

        with open(self._path, "w") as file:
            domtree.writexml(writer=file)
