# coding: utf-8
import json
from dataclasses import dataclass
from pathlib import Path
from typing import List

@dataclass
class AppConfig:
    
    lo_script_dir: List[str]
    """
    Path Like structure to libre office scripts director.
    """
    app_build_dir: List[str]
    """
    Path Like structure to build dir
    """
    app_res_dir: List[str]
    """
    Path Like structure to resources dir
    """
    app_res_blank_odt: List[str]
    """
    Path Like structure to resources dir.
    
    This is expected to be a subpath of ``app_res_dir``.
    """
    xml_manifest_namesapce: str
    """
    Name of LO manifest xml file such as `urn:oasis:names:tc:opendocument:xmlns:manifest:1.0`
    """


def read_config(config_file: str) -> AppConfig:
    """
    Gets config for given config file

    Args:
        config_file (str): Config file to load

    Returns:
        AppConfig: Configuration object
    """
    with open(config_file, 'r') as file:
        data = json.load(file)
        return AppConfig(**data)

def read_config_default() -> AppConfig:
    """
    Loads default configuration ``config.json``

    Returns:
        AppConfig: Configuration Object
    """
    root = Path(__file__).parent
    config_file = Path(root, 'config.json')
    return read_config(str(config_file))