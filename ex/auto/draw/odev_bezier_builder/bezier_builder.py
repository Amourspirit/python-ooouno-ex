from __future__ import annotations
from typing import List

import uno

from ooo.dyn.awt.point import Point
from ooo.dyn.drawing.polygon_flags import PolygonFlags

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.exceptions import ex as mEx
from ooodev.draw import Draw, DrawDoc, DrawPage
from ooodev.draw.shapes import OpenBezierShape
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr


class BezierBuilder:
    def __init__(self, fnm_point: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm_point, True)
        self._fnm_point = FileIO.get_absolute_path(fnm_point)

    def show(self) -> None:
        # with Lo.Loader(Lo.ConnectPipe()) as loader:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = DrawDoc(Draw.create_draw_doc(loader))
            slide = doc.slides[0]

            doc.set_visible()

            # self._draw_curve(slide) # same as bpts3.txt

            start_pt = Point()
            curve_pts: List[Point] = []
            is_open = self._read_points(
                fnm=self._fnm_point, start_pt=start_pt, curve_pts=curve_pts
            )
            _ = self._create_bezier(
                slide=slide, start_pt=start_pt, curve_pts=curve_pts, is_open=is_open
            )

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

    def _read_points(
        self, fnm: PathOrStr, start_pt: Point, curve_pts: List[Point]
    ) -> bool:
        is_open = True

        def process_line(line: str) -> None:
            nonlocal is_open
            ch = line[:1]
            if ch == "M":
                Lo.print("reading start point")
                pt = self._extract_point(line[1:].strip())
                start_pt.X = pt.X
                start_pt.Y = pt.Y
                Draw.print_point(start_pt)
            elif ch == "C":
                Lo.print("Reading curve points")
                self._set_curve(curve_pts=curve_pts, pts_str=line[1:].strip())
            elif ch == "Z":
                Lo.print("Read closed path flag")
                is_open = False
            # yield line

        with open(fnm, "r") as file:
            # strip of new line and remove anything after //
            # // for comment
            data = (row.partition("//")[0].rstrip() for row in file)
            # chain generator
            # remove empty lines
            data = (row for row in data if row)
            # each line should now start with M or C or Z or not recognized
            # remove all lines that do not match start Char
            data = (row for row in data if row[:1] in ("M", "C", "Z"))
            for row in data:
                process_line(row)

        return is_open

    def _draw_curve(self, slide: DrawPage[DrawDoc]) -> OpenBezierShape[DrawDoc]:
        # sample data, same as bpts3.txt
        path_pts: List[Point] = []
        path_flags: List[PolygonFlags] = []

        path_pts.append(Point(1_000, 2_500))
        path_flags.append(PolygonFlags.NORMAL)

        path_pts.append(Point(1_000, 1_000))  # control point
        path_flags.append(PolygonFlags.CONTROL)

        path_pts.append(Point(4_000, 1_000))  # control point
        path_flags.append(PolygonFlags.CONTROL)

        path_pts.append(Point(4_000, 2_500))
        path_flags.append(PolygonFlags.NORMAL)

        return slide.draw_bezier_open(pts=path_pts, flags=path_flags)

    def _get_integer(self, s: str) -> int:
        try:
            return int(s)
        except ValueError:
            return 0

    def _extract_point(self, pt_str: str) -> Point:
        # convert a string of the form "16400,10900" into a Point
        start_pt = Point()
        vals = [s.strip() for s in pt_str.split(",")]
        if len(vals) != 2:
            Lo.print("Could not parse the pont string into 2 parts")
        else:
            start_pt.X = self._get_integer(vals[0])
            start_pt.Y = self._get_integer(vals[1])
        return start_pt

    def _set_curve(self, curve_pts: List[Point], pts_str: str) -> None:
        # convert a string of the form "5400,14100 3600,3500 10000,7200" into points
        vals = [s.strip() for s in pts_str.split()]
        for val in vals:
            curve_pts.append(self._extract_point(val))

    def _create_bezier(
        self,
        slide: DrawPage[DrawDoc],
        start_pt: Point,
        curve_pts: List[Point],
        is_open: bool,
    ) -> OpenBezierShape[DrawDoc]:
        # The assumption is that the curve points use the C format from the SVG path
        # specification. Namely (x1 y1 x2 y2 x y)+,	where each tuple defines a cubic
        # Bézier curve. It starts at the startPt and ends at (x,y), using (x1,y1) as the
        # control point at the beginning of the curve and (x2,y2) as the control point
        # at the end of the curve.

        # The (x,y) becomes the new startPt for the next cubic curve.

        bez_pts: List[Point] = []
        flags: List[PolygonFlags] = []

        if len(curve_pts) % 3 != 0:
            raise mEx.NotSupportedError("Number of points must be a multiple of 3")

        path_step = int(len(curve_pts) / 3)
        for i in range(path_step):
            # fill in the points and flags for 1 step
            bez_pts.append(start_pt)
            flags.append(PolygonFlags.NORMAL)

            bez_pts.append(curve_pts[i * 3])  # (x1,y1) control point
            flags.append(PolygonFlags.CONTROL)

            bez_pts.append(curve_pts[i * 3 + 1])  # (x2,y2) control point
            flags.append(PolygonFlags.CONTROL)

            bez_pts.append(curve_pts[i * 3 + 2])  # (x,y) end point
            flags.append(PolygonFlags.CONTROL)
            start_pt = curve_pts[i * 3 + 2]  # and the next start point

        shape = slide.draw_bezier_open(pts=bez_pts, flags=flags)
        return shape
