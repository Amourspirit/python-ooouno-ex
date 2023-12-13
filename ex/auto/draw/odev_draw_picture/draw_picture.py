from __future__ import annotations

import uno

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import Draw, DrawDoc, DrawPage, Intensity, ZoomKind
from ooodev.utils.color import CommonColor
from ooodev.utils.lo import Lo


class DrawPicture:
    def show(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = DrawDoc(Draw.create_draw_doc(loader))
            doc.set_visible()
            Lo.delay(1_000)  # need delay or zoom may not occur
            doc.zoom(ZoomKind.ENTIRE_PAGE)

            curr_slide = doc.get_slide(idx=0)
            self._draw_shapes(curr_slide=curr_slide)

            s = curr_slide.draw_formula(
                formula="func e^{i %pi} + 1 = 0",
                x=70,
                y=20,
                width=75,
                height=40,
            )
            # Draw.report_pos_size(s)

            self._anim_shapes(curr_slide=curr_slide)

            s = curr_slide.find_shape_by_name("text1")
            Draw.report_pos_size(s.component)

            Lo.delay(2000)
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

    def _draw_shapes(self, curr_slide: DrawPage[DrawDoc]) -> None:
        line1 = curr_slide.draw_line(x1=50, y1=50, x2=200, y2=200)
        line1.component.LineColor = CommonColor.BLACK
        line1.set_dashed_line(is_dashed=True)

        # red ellipse; uses (x, y) width, height
        circle1 = curr_slide.draw_ellipse(x=100, y=100, width=75, height=25)
        circle1.component.FillColor = CommonColor.RED

        # rectangle with different fills; uses (x, y) width, height
        rect1 = curr_slide.draw_rectangle(x=70, y=100, width=75, height=25)
        rect1.component.FillColor = CommonColor.LIME

        text1 = curr_slide.draw_text(
            msg="Hello LibreOffice",
            x=120,
            y=120,
            width=60,
            height=30,
            font_size=24,
        )
        text1.component.Name = "text1"
        # Props.show_props("TextShape's Text Properties", Draw.get_text_properties(text1.component))

        # gray transparent circle; uses (x,y), radius
        circle2 = curr_slide.draw_circle(x=40, y=150, radius=20)
        circle2.component.FillColor = CommonColor.GRAY
        circle2.set_transparency(level=Intensity(25))

        # thick line; uses (x,y), angle clockwise from x-axis, length
        line2 = curr_slide.draw_polar_line(x=60, y=200, degrees=45, distance=100)
        line2.component.LineWidth = 300

    def _anim_shapes(self, curr_slide: DrawPage[DrawDoc]) -> None:
        # two animations of a circle and a line
        # he animation loop is:
        #    redraw shape, delay, update shape position/size

        # reduce circle size and move to the right
        xc = 40
        yc = 150
        radius = 40
        circle = None
        for _ in range(20):
            # move right
            if circle is not None:
                curr_slide.remove(circle.component)
            circle = curr_slide.draw_circle(x=xc, y=yc, radius=radius)

            Lo.delay(200)
            xc += 5
            radius *= 0.95

        x2 = 140
        y2 = 110
        line = None
        for _ in range(25):
            if line is not None:
                curr_slide.remove(line.component)
            line = curr_slide.draw_line(x1=40, y1=100, x2=x2, y2=y2)
            x2 -= 4
            y2 -= 4
