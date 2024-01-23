from __future__ import annotations
from typing import Any, List, Tuple, TYPE_CHECKING

import uno

from ooo.dyn.awt.point import Point
from ooo.dyn.drawing.polygon_flags import PolygonFlags

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import Draw, DrawDoc, DrawPage
from ooodev.draw.shapes import ClosedBezierShape
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr
from ooodev.format.draw.direct.area import Gradient, PresetGradientKind

if TYPE_CHECKING:
    from ooodev.draw.filter.export_png import ExportPngT
    from ooodev.draw.filter.export_jpg import ExportJpgT
    from ooodev.events.args.event_args_export import EventArgsExport

# https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Bezier_Shapes


class BezierBuilder:
    def __init__(self, fnm: PathOrStr) -> None:
        self._fnm = FileIO.get_absolute_path(fnm)

    def export(self, resolution: int = 96) -> None:
        # with Lo.Loader(Lo.ConnectPipe()) as loader:
        def on_exported_png(source: Any, args: EventArgsExport[ExportPngT]) -> None:
            self.on_exported()

        def on_exported_jpg(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
            self.on_exported()

        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            # create Draw slide
            doc = DrawDoc(Draw.create_draw_doc(loader))
            slide = doc.slides[0]

            doc.set_visible()

            ext = FileIO.get_ext(self._fnm).lower()
            shape = self._create_bezier(slide=slide)
            print(shape)

            # shape can also be accessed by index
            # the following line get the shape as ClosedBezierShape instance
            # shape = slide[0]

            shape.subscribe_event_shape_jpg_exported(on_exported_jpg)
            shape.subscribe_event_shape_png_exported(on_exported_png)
            if ext == "png":
                shape.export_shape_png(self._fnm, resolution=resolution)
            else:
                shape.export_shape_jpg(self._fnm, resolution=resolution)

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
        except Exception:
            Lo.close_office()
            raise

    def on_exported(self) -> None:
        msg = f"Shape exported to {self._fnm}"
        _ = MsgBox.msgbox(msg, "Exported", boxtype=MessageBoxType.INFOBOX)
        print(msg)

    def _create_bezier(self, slide: DrawPage[DrawDoc]) -> ClosedBezierShape:
        # The following values are exemplary and provokes that a PolyPolygon of
        # sixteen single polygons containing four points each is created. The
        # PolyPolygon total point count will be 64.
        # If control points are used they are allowed to appear as pair only,
        # before and after such pair has to be a normal point.
        # A bezier point sequence may look like
        # this (n=normal, c=control) : n c c n c c n n c c n

        polygon_count = 16
        width = 10_000
        height = 10_000

        point_grid: List[Tuple[Point, ...]] = []
        flags_grid: List[Tuple[PolygonFlags, ...]] = []

        ny = 0
        for _ in range(polygon_count):
            points: Tuple[Point, ...] = (
                Point(X=0, Y=ny),
                Point(X=width // 2, Y=height),
                Point(X=width // 2, Y=height),
                Point(X=width, Y=ny),
            )
            flags: Tuple[PolygonFlags, ...] = (
                PolygonFlags.NORMAL,
                PolygonFlags.CONTROL,
                PolygonFlags.CONTROL,
                PolygonFlags.NORMAL,
            )
            point_grid.append(points)
            flags_grid.append(flags)
            ny += height // polygon_count

        shape = slide.draw_bezier_closed(pts=point_grid, flags=flags_grid)
        # add a gradient to the shape
        shape.apply_styles(Gradient.from_preset(PresetGradientKind.MAHOGANY))
        return shape
