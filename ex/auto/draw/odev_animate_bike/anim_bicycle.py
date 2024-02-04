from __future__ import annotations

import uno

# from ooodev.utils.info import Info
from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import Draw, DrawDoc, DrawPage, PolySides
from ooodev.loader import Lo
from ooodev.utils.file_io import FileIO
from ooodev.utils.color import CommonColor
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.drawing.circle_kind import CircleKind


class AnimBicycle:
    def __init__(self, fnm_bike: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm_bike, raise_err=True)
        self._fnm_bike = FileIO.get_absolute_path(fnm_bike)

    def animate(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = DrawDoc.create_doc(loader=loader, visible=True)

            slide = doc.get_slide(idx=0)

            slide_size = slide.get_size_mm()

            square = slide.draw_polygon(x=125, y=125, sides=PolySides(4), radius=25)
            # square.component is com.sun.star.drawing.PolyPolygonShape service.
            square.component.FillColor = CommonColor.LIGHT_GREEN

            # default radius of 20, no. of sides
            pentagon = slide.draw_polygon(x=150, y=75, sides=PolySides(5))
            # pentagon.component is com.sun.star.drawing.PolyPolygonShape service.
            pentagon.component.FillColor = CommonColor.PURPLE

            xs = (10, 30, 10, 30)
            ys = (10, 100, 100, 10)

            slide.draw_lines(xs=xs, ys=ys)

            pie = slide.draw_ellipse(
                x=30, y=slide_size.Width - 100, width=40, height=20
            )
            pie.set_property(
                FillColor=CommonColor.LIGHT_SKY_BLUE,
                CircleStartAngle=9_000,  #   90 degrees ccw
                CircleEndAngle=36_000,  #    360 degrees ccw
                CircleKind=CircleKind.SECTION,
            )

            self._animate_bike(slide=slide)
            Lo.delay(2000)
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

    def _animate_bike(self, slide: DrawPage[DrawDoc]) -> None:
        shape = slide.draw_image(fnm=self._fnm_bike, x=60, y=100, width=90, height=50)

        pt = shape.get_position_mm()
        angle = shape.get_rotation()
        print(f"Start Angle: {int(angle)}")
        for i in range(19):
            shape.set_position(x=pt.X + (i * 5), y=pt.Y)  # move right
            shape.set_rotation(angle=angle + (i * 5))  # rotates ccw
            Lo.delay(200)

        print(f"Final Angle: {int(shape.get_rotation())}")
        Draw.print_matrix(shape.get_transformation())
