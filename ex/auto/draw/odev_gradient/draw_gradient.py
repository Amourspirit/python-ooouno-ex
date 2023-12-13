from __future__ import annotations
from enum import Enum

import uno
from ooo.dyn.awt.gradient import Gradient as Gradient

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import (
    Draw,
    DrawDoc,
    DrawPage,
    Angle,
    DrawingGradientKind,
    DrawingHatchingKind,
    DrawingBitmapKind,
    ZoomKind,
)
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class GradientKind(str, Enum):
    FILL = "fill"
    GRADIENT = "gradient"
    GRADIENT_NAME = "name"
    GRADIENT_NAME_PROPS = "name_props"
    HATCHING = "hatch"
    BITMAP = "bitmap"
    BITMAP_FILE = "file"


class DrawGradient:
    def __init__(
        self, gradient_kind: GradientKind, gradient_fnm: PathOrStr = ""
    ) -> None:
        self._gradient_kind = gradient_kind

        self._gradient_fnm = gradient_fnm
        if self._gradient_kind == GradientKind.BITMAP_FILE:
            # file has to be valid when bitmap file
            _ = FileIO.is_exist_file(self._gradient_fnm, True)
        self._x = 93
        self._y = 100
        self._width = 30
        self._height = 60
        self._start_color = CommonColor.LIME
        self._end_color = CommonColor.RED
        self._angle = 0
        self._name_gradient = DrawingGradientKind.NEON_LIGHT
        self._hatch_gradient = DrawingHatchingKind.GREEN_30_DEGREES
        self._bitmap_gradient = DrawingBitmapKind.FLORAL

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = DrawDoc(Draw.create_draw_doc(loader))
            doc.set_visible()
            Lo.delay(1_000)  # need delay or zoom may not occur
            doc.zoom(ZoomKind.ENTIRE_PAGE)

            curr_slide = doc.get_slide(idx=0)

            if self._gradient_kind == GradientKind.FILL:
                self._gradient_fill(curr_slide)
            elif self._gradient_kind == GradientKind.GRADIENT:
                self._gradient(curr_slide)
            elif self._gradient_kind == GradientKind.GRADIENT_NAME:
                self._gradient_name(curr_slide, False)
            elif self._gradient_kind == GradientKind.GRADIENT_NAME_PROPS:
                self._gradient_name(curr_slide, True)
            elif self._gradient_kind == GradientKind.HATCHING:
                self._gradient_hatching(curr_slide)
            elif self._gradient_kind == GradientKind.BITMAP:
                self._gradient_bitmap(curr_slide)
            elif self._gradient_kind == GradientKind.BITMAP_FILE:
                self._gradient_bitmap_file(curr_slide)

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

    def _gradient_fill(self, curr_slide: DrawPage[DrawDoc]) -> None:
        # rect1.component is com.sun.star.drawing.RectangleShape service which also implements com.sun.star.drawing.FillProperties service
        rect1 = curr_slide.draw_rectangle(
            x=self._x,
            y=self._y,
            width=self._width,
            height=self._height,
        )
        rect1.component.FillColor = self._start_color
        # other properties can be set
        # rect1.component.FillTransparence = 55

    def _gradient(self, curr_slide: DrawPage[DrawDoc]) -> None:
        rect1 = curr_slide.draw_rectangle(
            x=self._x,
            y=self._y,
            width=self._width,
            height=self._height,
        )
        rect1.set_gradient_color(
            start_color=self._start_color,
            end_color=self._end_color,
            angle=Angle(self._angle),
        )

    def _gradient_name(self, curr_slide: DrawPage[DrawDoc], set_props: bool) -> None:
        # rect1.component is com.sun.star.drawing.RectangleShape service which also implements com.sun.star.drawing.FillProperties service

        rect1 = curr_slide.draw_rectangle(
            x=self._x,
            y=self._y,
            width=self._width,
            height=self._height,
        )
        grad = rect1.set_gradient_color(name=self._name_gradient)
        if set_props:
            # grad = cast("Gradient", Props.get(rect1, "FillGradient"))
            # print(grad)
            grad.Angle = self._angle * 10  # in 1/10 degree units
            grad.StartColor = self._start_color
            grad.EndColor = self._end_color
            rect1.set_gradient_properties(grad=grad)

    def _gradient_hatching(self, curr_slide: DrawPage[DrawDoc]) -> None:
        rect1 = curr_slide.draw_rectangle(
            x=self._x,
            y=self._y,
            width=self._width,
            height=self._height,
        )
        rect1.set_hatch_color(name=self._hatch_gradient)

    def _gradient_bitmap(self, curr_slide: DrawPage[DrawDoc]) -> None:
        rect1 = curr_slide.draw_rectangle(
            x=self._x,
            y=self._y,
            width=self._width,
            height=self._height,
        )
        rect1.set_bitmap_color(name=self._bitmap_gradient)

    def _gradient_bitmap_file(self, curr_slide: DrawPage[DrawDoc]) -> None:
        rect1 = curr_slide.draw_rectangle(
            x=self._x,
            y=self._y,
            width=self._width,
            height=self._height,
        )
        rect1.set_bitmap_file_color(fnm=self._gradient_fnm)

    # region properties
    @property
    def angle(self) -> int:
        """Specifies angle"""
        return self._angle

    @angle.setter
    def angle(self, value: int):
        self._angle = value

    @property
    def x(self) -> int:
        """Specifies x"""
        return self._x

    @x.setter
    def x(self, value: int):
        self._x = value

    @property
    def y(self) -> int:
        """Specifies x"""
        return self._y

    @y.setter
    def y(self, value: int):
        self._y = value

    @property
    def width(self) -> int:
        """Specifies width"""
        return self._width

    @width.setter
    def width(self, value: int):
        self._width = value

    @property
    def height(self) -> int:
        """Specifies height"""
        return self._height

    @height.setter
    def height(self, value: int):
        self._height = value

    @property
    def start_color(self) -> int:
        """Specifies start_color"""
        return self._start_color

    @start_color.setter
    def start_color(self, value: int):
        self._start_color = value

    @property
    def end_color(self) -> int:
        """Specifies end_color"""
        return self._end_color

    @end_color.setter
    def end_color(self, value: int):
        self._end_color = value

    @property
    def name_gradient(self) -> DrawingGradientKind:
        """Specifies name_gradient"""
        return self._name_gradient

    @name_gradient.setter
    def name_gradient(self, value: DrawingGradientKind):
        self._name_gradient = value

    @property
    def hatch_gradient(self) -> DrawingHatchingKind:
        """Specifies hatch_gradient"""
        return self._hatch_gradient

    @hatch_gradient.setter
    def hatch_gradient(self, value: DrawingHatchingKind):
        self._hatch_gradient = value

    @property
    def bitmap_gradient(self) -> DrawingBitmapKind:
        """Specifies bitmap_gradient"""
        return self._bitmap_gradient

    @bitmap_gradient.setter
    def bitmap_gradient(self, value: DrawingBitmapKind):
        self._bitmap_gradient = value

    # endregion properties
