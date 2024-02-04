#!/usr/bin/env python
# coding: utf-8
import sys
import argparse
from typing import List
from pathlib import Path

import uno

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.format.writer.direct.char.font import Font
from ooodev.format.writer.direct.para.indent_space import Spacing, LineSpacing, ModeKind
from ooodev.loader import Lo
from ooodev.write import WriteDoc, WriteTextCursor, FamilyNamesKind
from ooodev.units import UnitMM
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.info import Info
from ooodev.format.writer.style.para import Para as StylePara, StyleParaKind
from ooodev.utils.color import CommonColor
from ooodev.format.writer.direct.page.header import Header
from ooodev.format.writer.direct.page.header.area import Img, PresetImageKind


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument(
        "-s",
        "--show",
        help="Show Document",
        action="store_true",
        dest="show",
        default=False,
    )
    parser.add_argument(
        "-v",
        "--verbose",
        help="Verbose output",
        action="store_true",
        dest="verbose",
        default=False,
    )


def read_text(fnm: Path, cursor: WriteTextCursor) -> None:
    sb: List[str] = []
    with open(fnm, "r") as file:
        i = 0
        for ln in file:
            line = ln.rstrip()  # remove new line \n
            if len(line) == 0:
                if len(sb) > 0:
                    cursor.append_para(" ".join(sb))
                sb.clear()
            elif line.startswith("Title: "):
                cursor.append_para(line[7:], styles=[StylePara(StyleParaKind.TITLE)])
            elif line.startswith("Author: "):
                cursor.append_para(line[8:], styles=[StylePara(StyleParaKind.SUBTITLE)])
            elif line.startswith("Part "):
                cursor.append_para(line, styles=[StylePara(StyleParaKind.HEADING_1)])
            else:
                sb.append(line)
            i += 1
            # if i > 20:
            #     break
        if len(sb) > 0:
            cursor.append_para(" ".join(sb))


def create_para_style(doc: WriteDoc, style_name: str) -> bool:
    try:
        # font 12 pt
        font = Font(name=Info.get_font_general_name(), size=12.0)

        # spacing below paragraphs
        spc = Spacing(below=UnitMM(4))

        # paragraph line spacing
        ln_spc = LineSpacing(mode=ModeKind.FIXED, value=UnitMM(6))

        _ = doc.create_style_para(style_name=style_name, styles=[font, spc, ln_spc])

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
        # parser.print_help()
        # return 0
        if len(sys.argv) == 1:
            pth = Path(__file__).parent / "data" / "scandal.txt"
            sys.argv.append("--show")
            sys.argv.append("-f")
            sys.argv.append(str(pth))

    # read the current command line args
    args = parser.parse_args()

    visible = args.show
    if visible:
        delay = 4_000
    else:
        delay = 0

    loader = Lo.load_office(
        connector=Lo.ConnectSocket(), opt=Lo.Options(verbose=args.verbose)
    )

    fnm = Path(args.file_path)

    try:
        doc = WriteDoc.create_doc(loader=loader, visible=visible)

        if not create_para_style(doc, "adParagraph"):
            raise RuntimeError("Could not create new paragraph style")

        styles = doc.get_style_names(FamilyNamesKind.PARAGRAPH_STYLES)
        print("Paragraph Styles")
        Lo.print_names(styles)

        text_range = doc.component.getText().getStart()
        # Load the paragraph style and apply it to the text range.
        para_style = StylePara("adParagraph")
        para_style.apply(text_range)

        # header formatting
        # create a header font style with a size of 9 pt, italic and dark green.
        header_font = Font(
            name=Info.get_font_general_name(),
            size=9.0,
            i=True,
            color=CommonColor.DARK_GREEN,
        )
        header_format = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=13.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
        )
        # create a header image from a preset
        header_img = Img.from_preset(PresetImageKind.MARBLE)
        # Set header can be passed a list of styles to format the header.
        doc.set_header(
            text=f"From: {fnm.name}",
            styles=[header_font, header_format, header_img],
        )

        # page format A4
        _ = doc.set_a4_page_format()
        # alternative setting of page format
        # from ooodev.format.writer.modify.page.page import PaperFormat, PaperFormatKind
        # page_size_style = PaperFormat.from_preset(preset=PaperFormatKind.A4)
        # page_size_style.apply(doc)

        _ = doc.set_page_numbers()

        cursor = doc.get_cursor()

        read_text(fnm=fnm, cursor=cursor)
        cursor.end_paragraph()

        cursor.append_para(f"Timestamp: {DateUtil.time_stamp()}")

        Lo.delay(delay)
        pth = Path.cwd() / "tmp"
        pth.mkdir(exist_ok=True)
        doc_pth = pth / "bigStory.doc"
        if visible:
            msg_result = doc.msgbox(
                "Do you wish to save document?",
                "Save",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.save_doc(fnm=doc_pth)
                print(f'Saved to: "{doc_pth}"')
        else:
            doc.save_doc(fnm=doc_pth)
            print(f'Saved to: "{doc_pth}"')
        if visible:
            msg_result = doc.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close_doc()
                Lo.close_office()
            else:
                print("Keeping document open")
        else:
            doc.close_doc()
            Lo.close_office()
    except Exception as e:
        Lo.print(e)
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
