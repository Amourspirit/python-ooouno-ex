from __future__ import annotations
import os
from pathlib import Path

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.format.writer.direct.char.font import Font
from ooodev.format.writer.direct.char.hyperlink import Hyperlink, TargetKind
from ooodev.format.writer.direct.frame.area import Color as FrameColor
from ooodev.format.writer.direct.frame.borders import (
    Side,
    Sides,
    BorderLineKind,
    LineSize,
)
from ooodev.format.writer.direct.para.alignment import Alignment
from ooodev.format.writer.direct.para.area import Color as ParaBgColor
from ooodev.format.writer.direct.para.outline_list import ListStyle, StyleListKind
from ooodev.format.writer.style.para import Para as ParaStyle
from ooodev.loader import Lo
from ooodev.write import WriteDoc
from ooodev.theme import ThemeGeneral
from ooodev.units import UnitMM
from ooodev.utils import color as color_util
from ooodev.utils.color import CommonColor
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.info import Info
from ooodev.utils.props import Props


def main() -> int:
    delay = 2_000  # delay so users can see changes.
    im_fnm = Path(__file__).parent / "data" / "skinner.png"

    loader = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = WriteDoc.create_doc(loader=loader, visible=True)

        cursor = doc.get_cursor()

        Props.show_obj_props(prop_kind="Cursor", obj=cursor.component)
        cursor.append("Some examples of simple text ")

        cursor.append_line(text="styles.", styles=[Font(b=True)])

        cursor.append_para(
            text="This line is written in red italics.",
            styles=[Font(color=CommonColor.DARK_RED).bold.italic],
        )

        cursor.append_para("Back to old style")
        cursor.append_line()
        cursor.append_para(text="A Nice Big Heading", styles=[ParaStyle().h1])
        cursor.append_para("The following points are important:")

        list_style = ListStyle(list_style=StyleListKind.NUM_123, num_start=-2)
        list_style.apply(cursor.component)

        cursor.append_para("Have a good breakfast")
        cursor.append_para("Have a good lunch")
        cursor.append_para("Have a good dinner")

        # Reset to default which set cursor to No List Style
        list_style.default.apply(cursor.component)
        cursor.end_paragraph()

        cursor.append_para("Breakfast should include:")

        # set cursor style to Number abc
        list_style = ListStyle(list_style=StyleListKind.NUM_abc, num_start=-2)
        list_style.apply(cursor.component)

        cursor.append_para("Porridge")
        cursor.append_para("Orange Juice")
        cursor.append_para("A Cup of Tea")
        # reset cursor number style
        list_style.default.apply(cursor.component)

        cursor.end_paragraph()

        cursor.append("This line ends with a bookmark.")
        cursor.add_bookmark("ad-bookmark")
        cursor.append_line()

        cursor.append_para("Here's some code:")

        tvc = doc.get_view_cursor()
        tvc.goto_range(cursor.component.getEnd(), False)

        y_pos = tvc.get_position().Y

        cursor.end_paragraph()
        code_font = Font(name=Info.get_font_mono_name(), size=10)
        code_font.apply(cursor.component)

        cursor.append_line("public class Hello")
        cursor.append_line("{")
        cursor.append_line("  public static void main(String args[]")
        cursor.append_line('  {  System.out.println("Hello World");  }')
        cursor.append_para("}  // end of Hello class")

        # reset the cursor formatting
        ParaStyle.default.apply(cursor.component)

        # Format the background color of the previous paragraph.
        bg_color = ParaBgColor(CommonColor.LIGHT_GRAY)
        cursor.style_prev_paragraph(styles=[bg_color])

        cursor.append_para("A text frame")

        pg = tvc.get_current_page()
        # add a text frame to the page and position it over the previous paragraph.
        # custom color is added via styles

        frame_color = FrameColor(CommonColor.DEFAULT_BLUE)
        # create a border
        bdr_sides = Sides(
            all=Side(
                line=BorderLineKind.SOLID, color=CommonColor.RED, width=LineSize.THIN
            )
        )

        _ = cursor.add_text_frame(
            text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
            ypos=y_pos,
            page_num=pg,
            width=UnitMM(40),
            height=UnitMM(15),
            styles=[frame_color, bdr_sides],
        )

        # Insert a hyperlink.
        cursor.append("A link to ")

        hl = Hyperlink(
            name="ODEV_GITHUB",
            url="https://github.com/Amourspirit/python_ooo_dev_tools",
            target=TargetKind.BLANK,
        )
        cursor.append("OOO Development Tools", styles=[hl])

        cursor.append_para(" Website.")

        # Add text based on theme color.
        # Supports LibreOffice 7.5 and above.
        if Info.version_info >= (7, 5, 0, 0):
            cursor.append_para()
            try:
                gen_theme_color = ThemeGeneral()
                if gen_theme_color.background_color < 0:
                    # -1 is Automatic color.
                    # assume light color
                    dark = False
                else:
                    rgb = color_util.RGB.from_int(gen_theme_color.background_color)
                    dark = rgb.is_dark()
            except Exception:
                dark = False
            if dark:
                cursor.append_para(
                    "This text is colored to match a dark theme.",
                    styles=[Font(color=color_util.StandardColor.BLUE_LIGHT3, size=16)],
                )
            else:
                cursor.append_para(
                    "This text is colored to match a light theme.",
                    styles=[Font(color=color_util.StandardColor.BLUE_DARK3, size=16)],
                )

        # start a new page
        cursor.page_break()

        # demonstrates how to lock the screen, Add content and then unlock the screen.
        with Lo.ControllerLock():
            # Lo.delay(delay)
            cursor.append_para(text="Image Example", styles=[ParaStyle().h2])

            cursor.append_para(f'The following image comes from "{im_fnm.name}":')
            cursor.end_paragraph()

            # For unknown reason if append is called with a new line here it cause a fatal error below on line 209 (Write.end_paragraph(cursor))
            # but only if image is add on line 208 (Write.add_image_shape(cursor=cursor, fnm=im_fnm)).
            cursor.append("Image as a link: ")

            img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
            cursor.add_image_link(
                fnm=im_fnm,
                width=img_size.width,
                height=img_size.height,
            )

            # enlarge by 1.5x
            h = round(img_size.height * 1.5)
            w = round(img_size.width * 1.5)

            cursor.add_image_link(fnm=im_fnm, width=w, height=h)
            cursor.end_paragraph()

        Lo.delay(delay)

        # Center previous paragraph
        cursor.style_prev_paragraph(styles=[Alignment().align_center])

        # check to see if we are on Linux
        if os.name != "posix" or Lo.bridge_connector.headless:
            # for some unknown reason when image shape is added in linux in GUI mode test will fail drastically.
            #   terminate called after throwing an instance of 'com::sun::star::lang::DisposedException'
            #   Fatal Python error: Aborted
            # on windows is fine. Running on linux in headless fine.

            cursor.append_line("Image as a shape: ")
            # add image as shape to page
            cursor.add_image_shape(fnm=im_fnm)
            cursor.end_paragraph()
            Lo.delay(delay)

        text_width = doc.get_page_text_width()

        cursor.add_line_divider(line_width=round(text_width * 0.5))

        # append timestamp as LO Fields.
        cursor.append_para("\nTimestamp: " + DateUtil.time_stamp() + "\n")
        cursor.append("Time (according to office): ")
        cursor.append_date_time()
        cursor.end_paragraph()

        # set some of the document properties.
        Info.set_doc_props(
            doc=doc.component,
            subject="Writer Text Example",
            title="Examples",
            author=":Barry-Thomas-Paul: Moss",
        )
        Lo.delay(delay)

        # move view cursor to bookmark position
        bookmark = doc.find_bookmark("ad-bookmark")
        assert bookmark is not None
        bm_range = bookmark.get_anchor()

        view_cursor = doc.get_view_cursor()
        view_cursor.goto_range(bm_range, False)

        Lo.delay(delay)
        msg_result = doc.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            pth = Path.cwd() / "tmp"
            pth.mkdir(parents=True, exist_ok=True)
            doc.save_doc(pth / "build.odt")

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
    except Exception:
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
