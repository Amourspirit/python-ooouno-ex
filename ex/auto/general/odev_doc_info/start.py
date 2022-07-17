#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import argparse
import sys

from ooodev.utils.lo import Lo
from ooodev.utils.info import Info
from ooodev.utils.file_io import FileIO
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext

# endregion Imports

# region args
def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-s",
        "--service",
        help="Show Services",
        action="store_true",
        dest="service",
        default=False,
    )
    parser.add_argument(
        "-i",
        "--interface",
        help="Show Interfaces",
        action="store_true",
        dest="interface",
        default=False,
    )
    parser.add_argument(
        "-x",
        "--xdoc",
        help="Show Methods for XTextDocument",
        action="store_true",
        dest="xdoc",
        default=False,
    )
    parser.add_argument(
        "-m",
        "--doc-meth",
        help="Show Methods for document ( may be long list )",
        action="store_true",
        dest="doc_meth",
        default=False,
    )
    parser.add_argument(
        "-p",
        "--property",
        help="Show properties",
        action="store_true",
        dest="property",
        default=False,
    )
    parser.add_argument(
        "-d",
        "--doc",
        help="Path to document to get infor for",
        action="store",
        dest="fnm_doc",
        required=True,
    )


# endregion args

# region Main
def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        parser.print_help(sys.stderr)
        return 1

    # read the current command line args
    args = parser.parse_args()

    with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=True))) as loader:
        fnm = args.fnm_doc
        doc_type = Info.get_doc_type(fnm=fnm)
        print(f"Doc type: {doc_type}")
        Props.show_doc_type_props(doc_type)

        try:
            doc = Lo.open_doc(fnm=fnm, loader=loader)
        except Exception:
            print(f"Could not open '{fnm}'")
            raise BreakContext.Break

        if args.service is True:
            print()
            print(" Services for this document: ".center(80, "-"))
            for service in Info.get_services(doc):
                print(f"  {service}")
            print()
            print(f"{Lo.Service.WRITER} is supported: {Info.is_doc_type(doc, Lo.Service.WRITER)}")
            print()

            print("  Available Services for this document: ".center(80, "-"))
            for i, service in enumerate(Info.get_available_services(doc)):
                print(f"  {service}")
            print(f"No. available services: {i}")

        if args.interface is True:
            print()
            print(" Interfaces for this document: ".center(80, "-"))
            for i, intfs in enumerate(Info.get_interfaces(doc)):
                print(f"  {intfs}")
            print(f"No. interfaces: {i}")

        if args.xdoc is True:
            print()
            print(f" Method for interface: com.sun.star.text.XTextDocument ".center(80, "-"))
            i = 0
            for meth in Info.get_methods("com.sun.star.text.XTextDocument"):
                print(f"  {meth}()")
                i+= 1
            print(f"No. methods: {i}")

        if args.property is True:
            print()
            print(" Properties for this document: ".center(80, "-"))
            i = 0
            for prop in Props.get_properties(doc):
                print(f"  {Props.show_property(prop)}")
                i += 1
            print(f"No. properties: {i}")

        if args.doc_meth is True:
            print()
            print(f" Method for entire document ".center(80, "-"))
            i = 0
            for i, meth in Info.get_methods_obj(doc):
                print(f"  {meth}()")
                i += 1
            print(f"No. methods: {i}")

        print()

        prop_name = "CharacterCount"
        print(f"Value of {prop_name}: {Props.get_property(doc, prop_name)}")

        Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# endregion main
