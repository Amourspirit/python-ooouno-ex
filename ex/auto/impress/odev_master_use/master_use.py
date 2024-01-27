# - list the shapes on the default master page.
#    - add shape and text to the master page
#    - set the footer text
#    - have normal slides use the slide number and footer on the master page
#
#    - create a second master page
#    - link one of the slides to the second master page
from __future__ import annotations

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.utils.lo import Lo
from ooodev.draw import Draw, ImpressDoc
from ooodev.utils.color import CommonColor


class MasterUse:
    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())
        try:
            doc = ImpressDoc.create_doc(loader)

            # report on the shapes on the default master page
            master_page = doc.get_master_page(idx=0)
            print("Default Master Page")
            Draw.show_shapes_info(master_page.component)

            # set the master page's footer text
            master_page.set_master_footer(text="Master Use Slides")

            # add a rectangle and text to the default master page
            # at the top-left of the slide
            sz = master_page.get_size_mm()
            _ = master_page.draw_rectangle(
                x=5,
                y=7,
                width=round(sz.Width / 6),
                height=round(sz.Height / 6),
            )
            _ = master_page.draw_text(
                msg="Default Master Page",
                x=10,
                y=15,
                width=100,
                height=10,
                font_size=24,
            )

            # set slide 1 to use the master page's slide number
            # but its own footer text
            slide1 = doc.slides[0]
            slide1.title_slide(title="Slide 1")

            # IsPageNumberVisible = True: use the master page's slide number
            # change the master page's footer for first slide; does not work if the master already has a footer
            slide1.set_property(
                IsPageNumberVisible=True,
                IsFooterVisible=True,
                FooterText="MU Slides",
            )

            # add three more slides, which use the master page's
            # slide number and footer
            for i in range(1, 4):  # 1, 2, 3
                slide = doc.insert_slide(idx=i)
                _ = slide.bullets_slide(title=f"Slide {i}")
                slide.set_property(IsPageNumberVisible=True, IsFooterVisible=True)

            # create master page 2
            master2 = doc.insert_master_page(idx=1)
            _ = master2.add_slide_number()

            print("Master Page 2")
            Draw.show_shapes_info(master2.component)

            # link master page 2 to third slide

            doc.slides[2].set_master_page(master2.component)

            # put ellipse and text on master page 2
            ellipse = master2.draw_ellipse(
                x=5,
                y=7,
                width=round(sz.Width / 6),
                height=round(sz.Height / 6),
            )
            ellipse.component.FillColor = CommonColor.GREEN_YELLOW
            _ = master2.draw_text(
                msg="Master Page 2",
                x=10,
                y=15,
                width=100,
                height=10,
                font_size=24,
            )

            doc.set_visible()

            Lo.delay(2_000)

            msg_result = MsgBox.msgbox(
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
