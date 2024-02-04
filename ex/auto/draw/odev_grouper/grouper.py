# region imports
from __future__ import annotations
from enum import Enum

import uno
from com.sun.star.container import XNameContainer
from com.sun.star.drawing import XShape
from com.sun.star.drawing import XShapeBinder
from com.sun.star.drawing import XShapeCombiner
from com.sun.star.drawing import XShapes

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import (
    Draw,
    DrawDoc,
    DrawPage,
    ShapeCombKind,
    DrawingShapeKind,
    GluePointsKind,
    GraphicStyleKind,
    ZoomKind,
)
from ooodev.draw.shapes import DrawShape
from ooodev.loader import Lo
from ooodev.utils.color import CommonColor
from ooodev.utils.info import Info
from ooodev.utils.kind.graphic_arrow_style_kind import GraphicArrowStyleKind
from ooodev.exceptions import ex as mEx

# endregion imports


# region Enums
class CombineEllipseKind(str, Enum):
    NONE = "none"
    GROUP = "group"
    BIND = "bind"
    COMBINE = "combine"


# endregion Enums


# region Class Grouper
class Grouper:
    # region constructor

    def __init__(self) -> None:
        self._overlap = False
        self._combine_kind = CombineEllipseKind.COMBINE

    # endregion constructor

    # region public methods

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = DrawDoc.create_doc(loader=loader, visible=True)
            Lo.delay(1_000)  # need delay or zoom may not occur
            doc.zoom(ZoomKind.ENTIRE_PAGE)

            curr_slide = doc.slides[0]

            print()
            print("Connecting rectangles ...")
            g_styles = Info.get_style_container(
                doc=doc.component, family_style_name="graphics"
            )
            # Info.show_container_names("Graphic styles", g_styles)

            self._connect_rectangles(slide=curr_slide, g_styles=g_styles)

            # create two ellipses
            slide_size = curr_slide.get_size_mm()
            width = 40
            height = 20
            x = round(((slide_size.width * 3) / 4) - (width / 2))
            y1 = 20
            if self.overlap:
                y2 = 30
            else:
                y2 = round((slide_size.height / 2) - (y1 + height))  # so separated

            s1 = curr_slide.draw_ellipse(x=x, y=y1, width=width, height=height)
            s2 = curr_slide.draw_ellipse(x=x, y=y2, width=width, height=height)

            Draw.show_shapes_info(curr_slide.component)

            # group, bind, or combine the ellipses
            print()
            print("Grouping (or binding) ellipses ...")
            if self._combine_kind == CombineEllipseKind.GROUP:
                self._group_ellipses(slide=curr_slide, s1=s1.component, s2=s2.component)
            elif self._combine_kind == CombineEllipseKind.BIND:
                self._bind_ellipses(slide=curr_slide, s1=s1.component, s2=s2.component)
            elif self._combine_kind == CombineEllipseKind.COMBINE:
                self._combine_ellipses(
                    slide=curr_slide, s1=s1.component, s2=s2.component
                )
            Draw.show_shapes_info(curr_slide.component)

            # combine some rectangles
            comp_shape = self._combine_rects(slide=curr_slide)
            Draw.show_shapes_info(curr_slide.component)

            print("Waiting a bit before splitting...")
            Lo.delay(3000)  # delay so user can see previous composition

            # split the combination into component shapes
            print()
            print("Splitting the combination ...")
            # split the combination into component shapes
            combiner = curr_slide.qi(XShapeCombiner, True)
            combiner.split(comp_shape.component)
            Draw.show_shapes_info(curr_slide.component)

            Lo.delay(1_500)
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

    # endregion public methods

    # region private methods

    def _connect_rectangles(
        self, slide: DrawPage[DrawDoc], g_styles: XNameContainer
    ) -> None:
        # draw two two labelled rectangles, one green, one blue, and
        #  connect them. Changing the connector to an arrow

        # dark green rectangle with shadow and text
        green_rect = slide.draw_rectangle(x=70, y=180, width=50, height=25)
        green_rect.component.FillColor = CommonColor.DARK_GREEN
        green_rect.component.Shadow = True
        green_rect.add_text(msg="Green Rect")

        # (blue, the default color) rectangle with shadow and text
        blue_rect = slide.draw_rectangle(x=140, y=220, width=50, height=25)
        blue_rect.component.Shadow = True
        blue_rect.add_text(msg="Blue Rect")

        # connect the two rectangles; from the first shape to the second
        conn_shape = slide.add_connector(
            shape1=green_rect.component,
            shape2=blue_rect.component,
            start_conn=GluePointsKind.BOTTOM,
            end_conn=GluePointsKind.TOP,
        )

        conn_shape.set_style(
            graphic_styles=g_styles,
            style_name=GraphicStyleKind.ARROW_LINE,
        )
        # arrow added at the 'from' end of the connector shape
        # and it thickens line and turns it black

        # use GraphicArrowStyleKind to lookup the values for LineStartName and LineEndName.
        # these are the the same names as seen in Draw, Graphic Sytles: Arrow Line dialog box.
        conn_shape.set_property(
            LineWidth=50,
            LineColor=CommonColor.DARK_ORANGE,
            LineStartName=str(GraphicArrowStyleKind.ARROW_SHORT),
            LineStartCenter=False,
            LineEndName=GraphicArrowStyleKind.NONE,
        )
        # Props.show_obj_props("Connector Shape", conn_shape)

        # report the glue points for the blue rectangle
        try:
            gps = blue_rect.get_glue_points()
            print("Glue Points for blue rectangle")
            for i, gp in enumerate(gps):
                print(f"  Glue point {i}: ({gp.Position.X}, {gp.Position.Y})")
        except mEx.DrawError:
            pass

    def _group_ellipses(self, slide: DrawPage[DrawDoc], s1: XShape, s2: XShape) -> None:
        shape_group = slide.add_shape(
            shape_type=DrawingShapeKind.GROUP_SHAPE,
            x=0,
            y=0,
            width=0,
            height=0,
        )
        shapes = shape_group.qi(XShapes, True)
        shapes.add(s1)
        shapes.add(s2)

    def _bind_ellipses(self, slide: DrawPage[DrawDoc], s1: XShape, s2: XShape) -> None:
        shapes = Lo.create_instance_mcf(
            XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
        )
        shapes.add(s1)
        shapes.add(s2)
        binder = slide.qi(XShapeBinder, True)
        binder.bind(shapes)

    def _combine_ellipses(
        self, slide: DrawPage[DrawDoc], s1: XShape, s2: XShape
    ) -> None:
        shapes = Lo.create_instance_mcf(
            XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
        )
        shapes.add(s1)
        shapes.add(s2)
        combiner = slide.qi(XShapeCombiner, True)
        combiner.combine(shapes)

    def _combine_rects(self, slide: DrawPage[DrawDoc]) -> DrawShape[DrawDoc]:
        print()
        print("Combining rectangles ...")
        r1 = slide.draw_rectangle(x=50, y=20, width=40, height=20)
        r2 = slide.draw_rectangle(x=70, y=25, width=40, height=20)
        shapes = Lo.create_instance_mcf(
            XShapes, "com.sun.star.drawing.ShapeCollection", raise_err=True
        )
        shapes.add(r1.component)
        shapes.add(r2.component)
        comb = slide.owner.combine_shape(
            shapes=shapes, combine_op=ShapeCombKind.COMBINE
        )
        return comb

    # endregion private methods

    # region Properties
    @property
    def overlap(self) -> bool:
        """Specifies if elipses are to overlap"""
        return self._overlap

    @overlap.setter
    def overlap(self, value: bool):
        self._overlap = value

    @property
    def combine_kind(self) -> CombineEllipseKind:
        """Specifies the kind of combining for ellipses"""
        return self._combine_kind

    @combine_kind.setter
    def combine_kind(self, value: CombineEllipseKind):
        self._combine_kind = value

    # endregion properties


# endregion Class Grouper
