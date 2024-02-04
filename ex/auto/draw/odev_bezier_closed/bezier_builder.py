from __future__ import annotations
from typing import List, Tuple

import uno

from ooo.dyn.awt.point import Point
from ooo.dyn.drawing.polygon_flags import PolygonFlags

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import DrawDoc, DrawPage
from ooodev.draw.shapes import ClosedBezierShape
from ooodev.loader import Lo
from ooodev.format.draw.direct.area import Gradient, PresetGradientKind

# https://wiki.openoffice.org/wiki/Documentation/DevGuide/Drawings/Bezier_Shapes


class BezierBuilder:
    def __init__(self) -> None:
        pass

    def show(self) -> None:
        # with Lo.Loader(Lo.ConnectPipe()) as loader:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            # create Draw slide
            doc = DrawDoc.create_doc(loader=loader, visible=True)
            slide = doc.slides[0]

            _ = self._create_bezier(slide=slide)

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
        except Exception:
            Lo.close_office()
            raise

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
