#!/usr/bin/env python
# coding: utf-8

from __future__ import annotations
import sys
import argparse
from typing import cast

import uno
from com.sun.star.text import XTextDocument

from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument("-s", "--show", help="Show Document", action="store_true", dest="show", default=False)
    parser.add_argument("-v", "--verbose", help="Verbose output", action="store_true", dest="verbose", default=False)


def show_styles(doc: XTextDocument) -> None:
    # get all the style families for this document
    style_families = Info.get_style_family_names(doc)
    print(f"No. of Style Family Names: {len(style_families)}")
    for style_family in style_families:
        print(f"  {style_family}")
    print()

    # list all the style names for each style family
    for i, style_family in enumerate(style_families):
        print(f'{i} "{style_family}" Style Family contains containers:')
        style_names = Info.get_style_names(doc, style_family)
        Lo.print_names(style_names)

    # Report the properties for the paragraph styles family under the "Standard" name
    Props.show_props('ParagraphStyles "Standard"', Info.get_style_props(doc, "ParagraphStyles", "Header"))
    print()

    # access other style families, other names...
    # Props.show_props('FrameStyles "Graphics"', Info.get_style_props(doc, "FrameStyles", "Graphics"))
    # print()

    # Props.show_props('NumberingStyles "List 1"', Info.get_style_props(doc, "NumberingStyles", "List 1"))
    # print()

    # Props.show_props('PageStyles "Envelope"', Info.get_style_props(doc, "PageStyles", "Standard"))
    # print()


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        parser.print_help()
        return 0

    # read the current command line args
    args = parser.parse_args()

    visible = args.show

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(
        Lo.Loader(connector=Lo.ConnectSocket(headless=not visible), opt=Lo.Options(verbose=args.verbose))
    ) as loader:

        fnm = cast(str, args.file_path)

        try:
            doc = Write.open_doc(fnm=fnm, loader=loader)
        except Exception as e:
            print(f"Could not open '{fnm}'")
            print(f"  {e}")
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            if visible:
                GUI.set_visible(is_visible=visible, odoc=doc)
            show_styles(doc)

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
