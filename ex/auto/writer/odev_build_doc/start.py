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
    # There are unnecessary delays throughout this code, just to better
    # illustrate this example in action.
    # (Underscores are allowed in numeric literals since Python 3.6.)
    delay = 2_000  # each delay, in ms

    image_filename = Path(__file__).parent / "data" / "skinner.png"

    # Normally, we'd use a `with ... as loader:` block here rather than a
    # `try ... except` block to ensure that `Lo.close_office()` gets called
    # automatically, but in this case we will offer the user the option of
    # keeping the document open in the LibreOffice app for further editing
    # after the completion of this script.
    loader = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = WriteDoc.create_doc(loader=loader, visible=True)
        cursor = doc.get_cursor()
        # Uncomment this to dump the cursor's properties to the console
        # Props.show_obj_props(prop_kind="Cursor", obj=cursor.component)

        # --------------------------------------------------------------------
        #                               First, a few straight-forward examples
        sample_paragraphs(cursor)
        sample_h1(cursor, "A Nice Big Heading")
        sample_lists(cursor)
        sample_bookmark(cursor, bookmark_name="ad-bookmark")
        sample_hyperlink(cursor)
        sample_text_conditional_on_theme(cursor)
        sample_horizontal_rule(doc, cursor)

        # --------------------------------------------------------------------
        #                    A code block with a text frame to the right of it
        cursor.append_para("Here's some code:")

        # We want to memorize the current vertical position, but a regular text
        # cursor doesn't know how to give us that. So, we need to create a View
        # Cursor and synchoronize it with the current position of the regular
        # text cursor first.
        text_view_cursor = doc.get_view_cursor()
        text_view_cursor.goto_range(cursor.component.getEnd(), False)
        y_pos_at_start_of_code_block = text_view_cursor.get_position().Y

        sample_code_block(cursor)

        sample_text_frame(cursor, text_view_cursor, y_pos_at_start_of_code_block)

        # --------------------------------------------------------------------
        #                                                          Some images
        # start a new page
        cursor.page_break()

        # this demonstrates how to lock the screen, add content and then unlock the screen.
        with Lo.ControllerLock():
            # Lo.delay(delay)
            sample_h2(cursor, "Image Example")

            img_size = sample_image(image_filename, cursor)

            sample_image_enlarged(image_filename, cursor, img_size, enlargement_factor=1.5)

        Lo.delay(delay)

        # Center previous paragraph
        cursor.style_prev_paragraph(styles=[Alignment().align_center])

        # check to see if we are on Linux
        if os.name != "posix" or Lo.bridge_connector.headless:
            # for some unknown reason when image shape is added in linux in GUI mode test will fail drastically.
            #   terminate called after throwing an instance of 'com::sun::star::lang::DisposedException'
            #   Fatal Python error: Aborted
            # on windows is fine. Running on linux in headless fine.

            sample_image_as_shape(image_filename, cursor)
            Lo.delay(delay)

        # --------------------------------------------------------------------
        #                                               Some metadata examples
        sample_horizontal_rule(doc, cursor)

        sample_datetime_insertions(cursor)

        set_some_document_properties(doc, cursor)

        # --------------------------------------------------------------------
        #                                          Return to previous bookmark
        Lo.delay(delay)
        return_to_previous_bookmark(doc, bookmark_name="ad-bookmark")

        # --------------------------------------------------------------------
        #                                               Converse with the User
        Lo.delay(delay)
        msg_result = doc.msgbox(
            "Do you wish to save document?",
            "Save",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            destination_folder = Path.cwd() / "tmp"
            destination_folder.mkdir(parents=True, exist_ok=True)
            doc.save_doc(destination_folder / "build.odt")
            print(f'Saved as {destination_folder}/build.odt')

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
            print("Keeping document open; therefore not calling Lo.close_office()")
    except Exception:
        Lo.close_office()
        raise

    return 0


# ############################################################################
#                                 THE SAMPLES
# ############################################################################


def sample_paragraphs(cursor):
    cursor.append("Some examples of simple text ")

    cursor.append_line(text="styles.", styles=[Font(b=True)])

    cursor.append_para(
        text="This line is written in red italics.",
        styles=[Font(color=CommonColor.DARK_RED).bold.italic],
    )

    cursor.append_para("Back to old style")
    cursor.append_line()


def sample_h1(cursor, text):
    cursor.append_para(text, styles=[ParaStyle().h1])


def sample_h2(cursor, text):
    cursor.append_para(text, styles=[ParaStyle().h2])


def sample_lists(cursor):
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

    # TODO what's the difference bewteen cursor.end_paragraph() and cursor.append_line()?
    cursor.end_paragraph()


def sample_bookmark(cursor, bookmark_name):
    cursor.append("This line ends with a bookmark.")
    cursor.add_bookmark(bookmark_name)
    cursor.append_line()


def sample_code_block(cursor):
    cursor.end_paragraph()  # insert newline
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


def sample_text_frame(cursor, text_view_cursor, y_pos_at_start_of_code_block):
    cursor.append_para("A text frame")

    pg = text_view_cursor.get_current_page()
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
        ypos=y_pos_at_start_of_code_block,
        page_num=pg,
        width=UnitMM(40),
        height=UnitMM(15),
        styles=[frame_color, bdr_sides],
    )


def sample_hyperlink(cursor):
    cursor.append("A link to ")

    hl = Hyperlink(
        name="ODEV_GITHUB",
        url="https://github.com/Amourspirit/python_ooo_dev_tools",
        target=TargetKind.BLANK,
    )
    cursor.append("OOO Development Tools", styles=[hl])

    cursor.append_para(" Website.")


def sample_text_conditional_on_theme(cursor):
    # This only works in LibreOffice 7.5 and above.
    if Info.version_info < (7, 5, 0, 0):
        return

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


def sample_image(image_filename, cursor):
    cursor.append_para(f'The following image comes from "{image_filename.name}":')
    cursor.end_paragraph()

    # For unknown reason if append is called with a new line here it cause a fatal error below on line 209 (Write.end_paragraph(cursor))
    # but only if image is add on line 208 (Write.add_image_shape(cursor=cursor, fnm=im_fnm)).
    cursor.append("Image as a link: ")

    img_size = ImagesLo.get_size_100mm(im_fnm=image_filename)
    cursor.add_image_link(
        fnm=image_filename,
        width=img_size.width,
        height=img_size.height,
    )

    return img_size


def sample_image_enlarged(image_filename, cursor, img_size, enlargement_factor=1.5):
    h = round(img_size.height * enlargement_factor)
    w = round(img_size.width * enlargement_factor)

    cursor.add_image_link(fnm=image_filename, width=w, height=h)
    cursor.end_paragraph()


def sample_image_as_shape(image_filename, cursor):
    cursor.append_line("Image as a shape: ")
    # add image as shape to page
    cursor.add_image_shape(fnm=image_filename)
    cursor.end_paragraph()


def sample_horizontal_rule(doc, cursor):
    text_width = doc.get_page_text_width()
    
    cursor.add_line_divider(line_width=round(text_width * 0.5))


def sample_datetime_insertions(cursor):
    # append a timestamp as text (static)
    cursor.append_para("\nTimestamp: " + DateUtil.time_stamp() + "\n")
    # append a timestamp as a LO Field (dynamic)
    cursor.append("Time (according to office): ")
    cursor.append_date_time()
    cursor.end_paragraph()


def set_some_document_properties(doc, cursor):
    # Convenience method for setting the three most common document prperties
    Info.set_doc_props(
        doc=doc.component,
        subject="Writer Text Example",
        title="Examples",
        author=":Barry-Thomas-Paul: Moss",
    )
    cursor.append_para("Three of the document's properties (title, subject, and author) have just been set.")
    # NOTE: Be careful that you pass in doc.component here, not just doc (even though the argument name is just doc).
    Info.print_doc_properties(doc.component)
    cursor.append_para("The document's properties have been dumped to the console for your inspection. "
                       "You can also inspect them via the operating system's file explorer.")


def return_to_previous_bookmark(doc, bookmark_name):
    bookmark = doc.find_bookmark(bookmark_name)
    assert bookmark is not None
    bm_range = bookmark.get_anchor()

    view_cursor = doc.get_view_cursor()
    view_cursor.goto_range(bm_range, False)


if __name__ == "__main__":
    raise SystemExit(main())
