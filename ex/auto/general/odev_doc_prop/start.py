#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import argparse
import platform

# debugging
import sys
from pathlib import Path
sys.path.insert(0, Path(__file__).parent.parent.parent.parent.parent)

from ooodev.utils.lo import Lo
from ooodev.utils.info import Info
from ooodev.utils.file_io import FileIO

# endregion Imports

# region args
def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-d",
        "--doc",
        help="Show Doc properties",
        action="store",
        dest="fnm_doc",
        required=True,
    )


# endregion args

# region Show


def show_filters() -> None:
    print()
    print(" File Filter Names for Office: ".center(50, "-"))
    filters = Info.get_filter_names()
    for filter in filters:
        print(f"  {filter}")
    print("-----------")
    print(f"No. of filters: {len(filter)}")


def show_services(title: str, srv_name: str | None = None) -> None:
    print()
    print(title.center(50, "-"))
    services = Info.get_service_names(srv_name)
    for service in services:
        print(f"  {service}")
    print("-----------")
    print(f"No. of services: {len(services)}")


# endregion Show

# region Main
def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
        doc = Lo.open_doc(fnm=args.fnm_doc, loader=loader)
        Info.print_doc_properties(doc)
        Info.set_doc_props(doc, "Example", "Examples", "Specific User")
        
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# endregion main
