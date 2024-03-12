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
from ooodev.utils.gui import GUI
from ooodev.adapter.drawing.rotation_descriptor_properties_partial import (
    RotationDescriptorPropertiesPartial,
)
from ooodev.units import Angle

from ooo.dyn.drawing.circle_kind import CircleKind


class RotateShape:
    def __init__(self) -> None:
        pass

    def animate(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = DrawDoc.create_doc(loader=loader, visible=True)

            slide = doc.slides[0]

            slide_size = slide.get_size_mm()

            square = slide.draw_rectangle(x=125, y=125, width=100, height=100)
            # square.component is com.sun.star.drawing.PolyPolygonShape service.
            square.component.FillColor = CommonColor.LIGHT_GREEN

            selector = doc.get_selection_supplier()
            selector.select(square.component)
            selection = selector.getSelection()
            assert selection is not None

            shapes = doc.get_selected_shapes()
            assert len(shapes) == 1
            shape = shapes[0]
            if isinstance(shape, RotationDescriptorPropertiesPartial):
                shape.rotate_angle += Angle(45)

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
