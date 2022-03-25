#!/usr/bin/env python
# coding: utf-8
"""
Creates system links for uno and uno_helper

see: docs/setup_env.rst
"""
from importlib.resources import path
import os
import sys
from pathlib import Path


def _get_virtual_path() -> str:
    spath = os.environ.get('VIRTUAL_ENV', None)
    if spath is not None:
        return spath
    return sys.base_exec_prefix

def add_links():
    ver = f"{sys.version_info[0]}.{sys.version_info[1]}"
    p_uno = Path('/usr/lib/python3/dist-packages/uno.py')
    p_uno_helper = Path('/usr/lib/python3/dist-packages/unohelper.py')
    if p_uno.exists():
        dest = Path(_get_virtual_path(),
                    f"lib/python{ver}/site-packages/uno.py")
        try:
            os.symlink(
                src=p_uno,
                dst=dest)
            print(f"Created system link: {p_uno} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")

    if p_uno_helper.exists():
        dest = Path(_get_virtual_path(),
                    f"lib/python{ver}/site-packages/unohelper.py")
        try:
            os.symlink(
                src=p_uno_helper,
                dst=dest)
            print(f"Created system link: {p_uno_helper} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")

def remove_links():
    ver = f"{sys.version_info[0]}.{sys.version_info[1]}"
    uno_path = Path(_get_virtual_path(),
                    f"lib/python{ver}/site-packages/uno.py")    
    if uno_path.exists():
        os.remove(uno_path)
        print("removed uno.py")
    else:
        print("uno.py does not exist in virtual env.")
    unohelper_path = Path(_get_virtual_path(),
                    f"lib/python{ver}/site-packages/unohelper.py")
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

if __name__ == '__main__':
    main()