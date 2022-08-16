#!/usr/bin/env python
# coding: utf-8
import sys
import argparse
from typing import List, Any
from pathlib import Path

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import LoEvents
from ooodev.office.write import Write
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext

from com.sun.star.beans import XPropertySet
from com.sun.star.style import XStyle
from com.sun.star.text import XTextCursor
from com.sun.star.text import XTextDocument

from ooo.dyn.style.line_spacing import LineSpacing
from ooo.dyn.style.line_spacing_mode import LineSpacingMode


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


def on_lo_print(source: Any, e: CancelEventArgs) -> None:
    # this method is a callback for ooodev internal printing
    # by setting e.canecl = True all internal printing of ooodev is suppressed
    e.cancel = True


def read_text(fnm: Path, cursor: XTextCursor) -> None:
    sb: List[str] = []
    with open(fnm, "r") as file:
        i = 0
        for ln in file:
            line = ln.rstrip()  # remove new line \n
            if len(line) == 0:
                if len(sb) > 0:
                    Write.append_para(cursor, " ".join(sb))
                sb.clear()
            elif line.startswith("Title: "):
                Write.append_para(cursor, line[7:])
                Write.style_prev_paragraph(cursor, "Title")
            elif line.startswith("Author: "):
                Write.append_para(cursor, line[8:])
                Write.style_prev_paragraph(cursor, "Subtitle")
            elif line.startswith("Part "):
                Write.append_para(cursor, line)
                Write.style_prev_paragraph(cursor, "Heading")
            else:
                sb.append(line)
            i += 1
            # if i > 20:
            #     break
        if len(sb) > 0:
            Write.append_para(cursor, " ".join(sb))


def create_para_style(doc: XTextDocument, style_name: str) -> bool:
    try:
        para_styles = Info.get_style_container(doc=doc, family_style_name="ParagraphStyles")

        # create new paragraph style properties set
        para_style = Lo.create_instance_msf(XStyle, "com.sun.star.style.ParagraphStyle", raise_err=True)
        props = Lo.qi(XPropertySet, para_style, raise_err=True)

        # set some properties
        props.setPropertyValue("CharFontName", Info.get_font_general_name())
        props.setPropertyValue("CharHeight", 12.0)
        props.setPropertyValue("ParaBottomMargin", 400)  # 4mm, in 100th mm

        line_spacing = LineSpacing(Mode=LineSpacingMode.FIX, Height=600)
        props.setPropertyValue("ParaLineSpacing", line_spacing)

        para_styles.insertByName(style_name, props)
        return True
    except Exception as e:
        print("Could not set paragraph style")
        print(f"  {e}")
    return False


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
    if visible:
        delay = 4_000
    else:
        delay = 0

    if not args.verbose:
        # hook ooodev internal printing event
        LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(Lo.ConnectSocket(headless=not visible))) as loader:

        fnm = Path(args.file_path)

        try:
            doc = Write.create_doc(loader=loader)
        except Exception as e:
            print(e)
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            if visible:
                GUI.set_visible(is_visible=visible, odoc=doc)

            styles = Info.get_style_names(doc, "ParagraphStyles")
            print("Paragraph Styles")
            Lo.print_names(styles)

            if not create_para_style(doc, "adParagraph"):
                print("Could not create new paragraph style")
                # office will close and with statement is exited
                raise BreakContext.Break

            xtext_range = doc.getText().getStart()
            Props.set_property(xtext_range, "ParaStyleName", "adParagraph")

            Write.set_header(text_doc=doc, text=f"From: {fnm.name}")
            Write.set_a4_page_format(doc)
            Write.set_page_numbers(doc)

            cursor = Write.get_cursor(doc)

            read_text(fnm=fnm, cursor=cursor)
            Write.end_paragraph(cursor)

            Write.append_para(cursor, f"Timestamp: {DateUtil.time_stamp()}")

            Lo.delay(delay)

            Write.save_doc(text_doc=doc, fnm="bigStory.doc")

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
