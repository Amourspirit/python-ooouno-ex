#!/usr/bin/env python
# coding: utf-8
"""
Creates system links for uno and uno_helper

see: docs/setup_env.rst
"""
import os
import sys
import shutil
from typing import Optional, Union
from pathlib import Path


def _get_virtual_path() -> str:
    spath = os.environ.get("VIRTUAL_ENV", None)
    if spath is not None:
        return spath
    return sys.base_exec_prefix


def _get_uno_path() -> Path:
    if os.name == "nt":

        p_uno = Path(os.environ["PROGRAMFILES"], "LibreOffice", "program")
        if p_uno.exists() is False or p_uno.is_dir() is False:
            p_uno = Path(os.environ["PROGRAMFILES(X86)"], "LibreOffice", "program")
        if not p_uno.exists():
            raise FileNotFoundError("Uno Source Dir not found.")
        if not p_uno.is_dir():
            raise NotADirectoryError("UNO source is not a Directory")
        return p_uno
    else:
        p_uno = Path("/usr/lib/python3/dist-packages")
        if not p_uno.exists():
            raise FileNotFoundError("Uno Source Dir not found.")
        if not p_uno.is_dir():
            raise NotADirectoryError("UNO source is not a Directory")
        return p_uno


def _get_env_site_packeges_dir() -> Union[Path, None]:
    v_path = _get_virtual_path()
    p_site = Path(v_path, "Lib", "site-packages")
    if p_site.exists() and p_site.is_dir():
        return p_site

    ver = f"{sys.version_info[0]}.{sys.version_info[1]}"
    p_site = Path(v_path, "lib", f"python{ver}", "site-packages")
    if p_site.exists() and p_site.is_dir():
        return p_site
    return None


def add_links(uno_src_dir: Optional[str] = None):
    if isinstance(uno_src_dir, str):
        str_cln = uno_src_dir.strip()
        if len(str_cln) == 0:
            p_uno_dir = _get_uno_path()
        else:
            p_uno_dir = Path(str_cln)
            if not p_uno_dir.exists():
                raise FileNotFoundError(f"Uno Source Dir not found: {uno_src_dir}")
            if not p_uno_dir.is_dir():
                raise NotADirectoryError(
                    f"UNO source is not a Directory: {uno_src_dir}"
                )
    else:
        p_uno_dir = _get_uno_path()
    p_site_dir = _get_env_site_packeges_dir()
    if p_site_dir is None:
        print("Unable to find site_packages direct in virtual enviornment")
        return

    p_uno = Path(p_uno_dir, "uno.py")
    p_uno_helper = Path(p_uno_dir, "unohelper.py")
    if p_uno.exists():
        dest = Path(p_site_dir, "uno.py")
        try:
            os.symlink(src=p_uno, dst=dest)
            print(f"Created system link: {p_uno} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")
        except OSError:
            # OSError: [WinError 1314] A required privilege is not held by the client
            print(f"Unable to create system link for  '{p_uno.name}'. Attempting copy.")
            shutil.copy2(p_uno, dest)
            print(f"Copied file: {p_uno} -> {dest}")

    if p_uno_helper.exists():
        dest = Path(p_site_dir, "unohelper.py")
        try:
            os.symlink(src=p_uno_helper, dst=dest)
            print(f"Created system link: {p_uno_helper} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")
        except OSError:
            # OSError: [WinError 1314] A required privilege is not held by the client
            print(
                f"Unable to create system link for  '{p_uno_helper.name}'. Attempting copy."
            )
            shutil.copy2(p_uno_helper, dest)
            print(f"Copied file: {p_uno_helper} -> {dest}")


def remove_links():
    p_site_dir = _get_env_site_packeges_dir()
    if p_site_dir is None:
        print("Unable to find site_packages direct in virtual enviornment")
        return

    uno_path = Path(p_site_dir, "uno.py")
    if uno_path.exists():
        os.remove(uno_path)
        print("removed uno.py")
    else:
        print("uno.py does not exist in virtual env.")
    unohelper_path = Path(p_site_dir, "unohelper.py")
    if unohelper_path.exists():
        os.remove(unohelper_path)
        print("removed unohelper.py")
    else:
        print("unohelper.py does not exist in virtual env.")


def main():
    if len(sys.argv) == 2:
        arg = sys.argv[1]
        if arg == "-r" or arg == "--remove":
            remove_links()
            return
        if arg == "-a" or arg == "-add":
            add_links()
            return
    print("for add links use -a or --add\nfor remove use -r or --remove")


if __name__ == "__main__":
    main()
