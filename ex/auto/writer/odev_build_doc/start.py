#!/usr/bin/env python
# coding: utf-8
from functools import partial

from ooodev.office.write import Write
from ooodev.utils.color import CommonColor
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.wrapper.break_context import BreakContext


def main() -> int:

    delay = 2_000  # delay so users can see changes.

    im_fnm = FileIO.get_absolute_path("../../../../resources/image/skinner.png")
    if not im_fnm.exists():
        print("resource image 'skinner.png' not found.")
        print("Unable to continue.")
        return 1

    # Using Lo.Loader context manager wraped by BreakContext load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/wrapper/break_context.html
    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with BreakContext(Lo.Loader(Lo.ConnectSocket())) as loader:

        try:
            doc = Write.create_doc(loader=loader)
        except Exception as e:
            print(e)
            # office will close and with statement is exited
            raise BreakContext.Break

        try:
            GUI.set_visible(is_visible=True, odoc=doc)

            cursor = Write.get_cursor(doc)

            # take advantage of a few partial functions
            append = partial(Write.append, cursor)
            para = partial(Write.append_para, cursor)
            nl = partial(Write.append_line, cursor)
            np = partial(Write.end_paragraph, cursor)
            get_pos = partial(Write.get_position, cursor)

            Props.show_obj_props(prop_kind="Cursor", obj=cursor)
            append(text="Some examples of simple text ")
            pos = get_pos()
            append("styles.")
            append(ctl_char=Write.ControlCharacter.LINE_BREAK)
            Write.style_left_bold(cursor=cursor, pos=pos)

            pos = get_pos()
            para("This line is written in red italics.")
            Write.style_left_color(cursor=cursor, pos=pos, color=CommonColor.DARK_RED)
            Write.style_left_italic(cursor=cursor, pos=pos)

            Write.append_para(cursor=cursor, text="Back to old style")
            nl()

            Write.append_para(cursor=cursor, text="A Nice Big Heading")
            Write.style_prev_paragraph(cursor, "Heading 1")

            Write.append_para(cursor, "The following points are important:")
            pos = get_pos()
            Write.append_para(cursor, "Have a good breakfast")
            Write.append_para(cursor, "Have a good lunch")
            Write.append_para(cursor, "Have a good dinner")

            Write.style_left(cursor, pos, "NumberingStyleName", "Numbering 123")

            tvc = Write.get_view_cursor(doc)

            np()
            para("Breakfast should include:")
            pos = get_pos()
            para("Porridge")
            para("Orange Juice")
            para("A Cup of Tea")
            Write.style_left(cursor, pos, "NumberingStyleName", "Numbering abc")
            np()

            append("This line ends with a bookmark.")
            Write.add_bookmark(cursor=cursor, name="ad-bookmark")
            para("\n")

            para("Here's some code:")

            tvc = Write.get_view_cursor(doc)
            tvc.gotoRange(cursor.getEnd(), False)

            ypos = tvc.getPosition().Y

            np()
            pos = get_pos()
            nl("public class Hello")
            nl("{")
            nl("  public static void main(String args[]")
            nl('  {  System.out.println("Hello World");  }')
            para("}  // end of Hello class")

            Write.style_left_code(cursor, pos)

            # It is nice to format the background color of the previous paragraph. However using the next line is buggy.
            # Write.style_prev_paragraph(cursor=cursor, prop_name="ParaBackColor", prop_val=CommonColor.LIGHT_GRAY)
            # There seems to be a bug with LO when setting ParaBackColor property using API.
            # After save the background color is lost.
            # See: https://forum.openoffice.org/en/forum/viewtopic.php?f=7&t=88722
            # Thankfully we can work around this by using dispatch commands
            #
            # create a property containing the background color desired
            bg_props = Props.make_props(BackgroundColor=CommonColor.LIGHT_GRAY)
            # Using a special Write method dispatch the background color from current position to previous postion
            Write.dispatch_cmd_left(vcursor=tvc, pos=pos, cmd="BackgroundColor", props=bg_props)

            # get the new current position, should be end of current page content
            # reset the style back to Document Default Paragraph Style so it is not affect by previous dispatch command.
            # make the default Properties
            default_props = Props.make_props(Template="Default Paragraph Style", Family=2)
            # dispatch using Lo.dispatch_cmd as it is not needed to apply to a range at this point
            Lo.dispatch_cmd(cmd="StyleApply", props=default_props)

            para("A text frame")

            pg = Write.get_current_page(tvc)
            # add a text frame to the page and position it over the previous paragraph.
            # custom colors can be added but in this case will stick with defaults.
            Write.add_text_frame(
                cursor=cursor,
                ypos=ypos,
                text="This is a newly created text frame.\nWhich is over on the right of the page, next to the code.",
                page_num=pg,
                width=4000,
                height=1500,
            )

            # Create text that contains a hyperlink
            append("A link to ")
            pos = get_pos()
            append("OOO Development Tools")
            url_str = "https://github.com/Amourspirit/python_ooo_dev_tools"
            Write.style_left(cursor=cursor, pos=pos, prop_name="HyperLinkURL", prop_val=url_str)
            append(" Website.")
            Write.end_paragraph(cursor)

            # start a new page
            Write.page_break(cursor)

            # demonstrates how to lock the screen, Add content and then unlock the screen.
            with Lo.ControllerLock():
                Lo.delay(delay)
                para("Image Example")
                Write.style_prev_paragraph(cursor, "Heading 2")

                para(f'The following image comes from "{im_fnm.name}":')
                np()

                append(f"Image as a link: ")

                img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
                Write.add_image_link(doc, cursor, im_fnm, img_size.Width, img_size.Height)

                # enlarge by 1.5x
                h = round(img_size.Height * 1.5)
                w = round(img_size.Width * 1.5)

                Write.add_image_link(doc, cursor, im_fnm, w, h)
                Write.end_paragraph(cursor)

            Lo.delay(delay)

            # Center previous paragraph
            Write.style_prev_paragraph(cursor=cursor, prop_name="ParaAdjust", prop_val=Write.ParagraphAdjust.CENTER)

            # add image as shape to page
            append("Image as a shape: ")
            Write.add_image_shape(cursor=cursor, fnm=im_fnm)
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
            Info.set_doc_props(
                doc=doc, subject="Writer Text Example", title="Examples", author=":Barry-Thomas-Paul: Moss"
            )
            Lo.delay(delay)

            # move view cursor to bookmark position
            bookmark = Write.find_bookmark(doc, "ad-bookmark")
            bm_range = bookmark.getAnchor()

            view_cursor = Write.get_view_cursor(doc)
            view_cursor.gotoRange(bm_range, False)

            Lo.delay(delay)
            Lo.save_doc(doc, "build.odt")

        finally:
            Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
