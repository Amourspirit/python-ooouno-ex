from __future__ import annotations
import os
from pathlib import Path
from functools import partial

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.format.writer.direct.char.font import Font
from ooodev.format.writer.direct.char.hyperlink import Hyperlink, TargetKind
from ooodev.format.writer.direct.frame.area import Color as FrameColor
from ooodev.format.writer.direct.frame.borders import Side, Sides, BorderLineKind, LineSize
from ooodev.format.writer.direct.para.alignment import Alignment
from ooodev.format.writer.direct.para.area import Color as ParaBgColor
from ooodev.format.writer.direct.para.outline_list import ListStyle, StyleListKind
from ooodev.format.writer.style.para import Para as ParaStyle
from ooodev.office.write import Write
from ooodev.theme import ThemeGeneral
from ooodev.units import UnitMM
from ooodev.utils import color as color_util
from ooodev.utils.color import CommonColor
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.gui import GUI
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props


def main() -> int:

    delay = 2_000  # delay so users can see changes.
    im_fnm = Path(__file__).parent / "data" / "skinner.png"

    loader = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = Write.create_doc(loader=loader)
        GUI.set_visible(visible=True, doc=doc)

        cursor = Write.get_cursor(doc)

        # take advantage of a few partial functions
        nl = partial(Write.append_line, cursor)
        np = partial(Write.end_paragraph, cursor)

        Props.show_obj_props(prop_kind="Cursor", obj=cursor)
        Write.append(cursor, "Some examples of simple text ")

        Write.append_line(cursor=cursor, text="styles.", styles=[Font(b=True)])

        Write.append_para(
            cursor=cursor,
            text="This line is written in red italics.",
            styles=[Font(color=CommonColor.DARK_RED).bold.italic],
        )

        Write.append_para(cursor=cursor, text="Back to old style")
        nl()

        Write.append_para(cursor=cursor, text="A Nice Big Heading", styles=[ParaStyle().h1])

        Write.append_para(cursor, "The following points are important:")

        list_style = ListStyle(list_style=StyleListKind.NUM_123, num_start=-2)
        list_style.apply(cursor)

        Write.append_para(cursor, "Have a good breakfast")
        Write.append_para(cursor, "Have a good lunch")
        Write.append_para(cursor, "Have a good dinner")

        # Reset to default which set cursor to No List Style
        list_style.default.apply(cursor)
        np()

        Write.append_para(cursor=cursor, text="Breakfast should include:")

        # set cursor style to Number abc
        list_style = ListStyle(list_style=StyleListKind.NUM_abc, num_start=-2)
        list_style.apply(cursor)

        Write.append_para(cursor, "Porridge")
        Write.append_para(cursor, "Orange Juice")
        Write.append_para(cursor, "A Cup of Tea")
        # reset cursor number style
        list_style.default.apply(cursor)

        np()

        Write.append(cursor, "This line ends with a bookmark.")
        Write.add_bookmark(cursor=cursor, name="ad-bookmark")
        Write.append_line(cursor)

        Write.append_para(cursor, "Here's some code:")
        tvc = Write.get_view_cursor(doc)

        tvc = Write.get_view_cursor(doc)
        tvc.gotoRange(cursor.getEnd(), False)

        y_pos = tvc.getPosition().Y

        np()
        code_font = Font(name=Info.get_font_mono_name(), size=10)
        code_font.apply(cursor)

        nl("public class Hello")
        nl("{")
        nl("  public static void main(String args[]")
        nl('  {  System.out.println("Hello World");  }')
        Write.append_para(cursor, "}  // end of Hello class")

        # reset the cursor formatting
        ParaStyle.default.apply(cursor)

        # Format the background color of the previous paragraph.
        bg_color = ParaBgColor(CommonColor.LIGHT_GRAY)
        Write.style_prev_paragraph(cursor=cursor, styles=[bg_color])

        Write.append_para(cursor, "A text frame")

        pg = Write.get_current_page(tvc)
        # add a text frame to the page and position it over the previous paragraph.
        # custom color is added via styles

        frame_color = FrameColor(CommonColor.DEFAULT_BLUE)
        # create a border
        bdr_sides= Sides(all=Side(line=BorderLineKind.SOLID, color=CommonColor.RED, width=LineSize.THIN))
        
        Write.add_text_frame(
            cursor=cursor,
            ypos=y_pos,
            text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
            page_num=pg,
            width=UnitMM(40),
            height=UnitMM(15),
            styles=[frame_color, bdr_sides],
        )

        # Insert a hyperlink.
        Write.append(cursor, "A link to ")

        hl = Hyperlink(
            name="ODEV_GITHUB", url="https://github.com/Amourspirit/python_ooo_dev_tools", target=TargetKind.BLANK
        )
        Write.append(cursor, "OOO Development Tools", styles=[hl])

        Write.append_para(cursor, " Website.")

        # Add text based on theme color.
        # Supports LibreOffice 7.5 and above.
        if Info.version_info >= (7, 5, 0, 0):
            Write.append_para(cursor)
            gen_theme_color = ThemeGeneral()
            if gen_theme_color.background_color < 0:
                # -1 is Automatic color.
                # assume light color
                dark = False
            else:
                rgb = color_util.RGB.from_int(gen_theme_color.background_color)
                dark = rgb.is_dark()
            if dark:
                Write.append_para(
                    cursor,
                    "This text is colored to match a dark theme.",
                    styles=[Font(color=color_util.StandardColor.BLUE_LIGHT3, size=16)],
                )
            else:
                Write.append_para(
                    cursor,
                    "This text is colored to match a light theme.",
                    styles=[Font(color=color_util.StandardColor.BLUE_DARK3, size=16)],
                )

        # start a new page
        Write.page_break(cursor)

        # demonstrates how to lock the screen, Add content and then unlock the screen.
        with Lo.ControllerLock():
            # Lo.delay(delay)
            Write.append_para(cursor=cursor, text="Image Example", styles=[ParaStyle().h2])

            Write.append_para(cursor, f'The following image comes from "{im_fnm.name}":')
            np()

            # For unknown reason if append is called with a new line here it cause a fatal error below on line 209 (Write.end_paragraph(cursor))
            # but only if image is add on line 208 (Write.add_image_shape(cursor=cursor, fnm=im_fnm)).
            Write.append(cursor, "Image as a link: ")

            img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
            Write.add_image_link(doc=doc, cursor=cursor, fnm=im_fnm, width=img_size.width, height=img_size.height)

            # enlarge by 1.5x
            h = round(img_size.height * 1.5)
            w = round(img_size.width * 1.5)

            Write.add_image_link(doc=doc, cursor=cursor, fnm=im_fnm, width=w, height=h)
            Write.end_paragraph(cursor)

        Lo.delay(delay)

        # Center previous paragraph
        Write.style_prev_paragraph(cursor=cursor, styles=[Alignment().align_center])


        # check to see if we are on Linux
        if os.name != "posix" or Lo.bridge_connector.headless:
            # for some unknown reason when image shape is added in linux in GUI mode test will fail drastically.
            #   terminate called after throwing an instance of 'com::sun::star::lang::DisposedException'
            #   Fatal Python error: Aborted
            # on windows is fine. Running on linux in headless fine.

            Write.append_line(cursor, "Image as a shape: ")
            # add image as shape to page
            # Write.add_image_shape(cursor=cursor, fnm=im_fnm)
            Write.end_paragraph(cursor)
            Lo.delay(delay)

        text_width = Write.get_page_text_width(doc)

        Write.add_line_divider(cursor=cursor, line_width=round(text_width * 0.5))

        # append timestamp as LO Fields.
        Write.append_para(cursor, "\nTimestamp: " + DateUtil.time_stamp() + "\n")
        Write.append(cursor, "Time (according to office): ")
        Write.append_date_time(cursor=cursor)
        Write.end_paragraph(cursor)

        # set some of the document properties.
        Info.set_doc_props(doc=doc, subject="Writer Text Example", title="Examples", author=":Barry-Thomas-Paul: Moss")
        Lo.delay(delay)

        # move view cursor to bookmark position
        bookmark = Write.find_bookmark(doc, "ad-bookmark")
        bm_range = bookmark.getAnchor()

        view_cursor = Write.get_view_cursor(doc)
        view_cursor.gotoRange(bm_range, False)

        Lo.delay(delay)
        msg_result = MsgBox.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Lo.save_doc(doc, "build.odt")

        msg_result = MsgBox.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Lo.close_doc(doc)
            Lo.close_office()
        else:
            print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
