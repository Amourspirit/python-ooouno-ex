from __future__ import annotations
from typing import cast

import uno

from ooodev.draw import (
    Draw,
    ImpressDoc,
    ImpressPage,
    DrawingBitmapKind,
    DrawingHatchingKind,
    ShapeDispatchKind,
    DrawingGradientKind,
    ImageOffset,
    ZoomKind,
)
from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.drawing.hatch import Hatch
from ooo.dyn.drawing.line_style import LineStyle
from ooo.dyn.presentation.animation_effect import AnimationEffect
from ooo.dyn.presentation.animation_speed import AnimationSpeed
from ooo.dyn.presentation.click_action import ClickAction

try:
    # only windows
    from odevgui_win.draw_dispatcher import DrawDispatcher
except ImportError:
    DrawDispatcher = None


class MakeSlides:
    def __init__(
        self, fnm_wildlife: PathOrStr, fnm_clock: PathOrStr, fnm_img: PathOrStr
    ) -> None:
        _ = FileIO.is_exist_file(fnm=fnm_wildlife, raise_err=True)
        _ = FileIO.is_exist_file(fnm=fnm_clock, raise_err=True)
        _ = FileIO.is_exist_file(fnm=fnm_img, raise_err=True)
        self._fnm_wildlife = FileIO.get_absolute_path(fnm_wildlife)
        self._fnm_clock = FileIO.get_absolute_path(fnm_clock)
        self._fnm_img = FileIO.get_absolute_path(fnm_img)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = ImpressDoc.create_doc(loader)
            curr_slide = doc.slides[0]

            doc.set_visible()
            Lo.delay(500)  # delay to make sure zoom takes
            doc.zoom(ZoomKind.ENTIRE_PAGE)

            curr_slide.title_slide(
                title="Python-Generated Slides",
                sub_title="Using LibreOffice",
            )

            # second slide
            curr_slide = doc.add_slide()
            self._do_bullets(curr_slide=curr_slide)

            # third slide: title and video
            curr_slide = doc.add_slide()
            curr_slide.title_only_slide("Clock Video")
            curr_slide.draw_media(fnm=self._fnm_clock, x=20, y=70, width=50, height=50)

            # fourth slide
            curr_slide = doc.add_slide()
            self._button_shapes(curr_slide=curr_slide)

            # fifth slide
            if DrawDispatcher:
                # windows only
                # a bit slow due to gui interaction but a good demo
                self._dispatch_shapes(doc)

            Lo.print(f"Total no. of slides: {doc.get_slides_count()}")

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

    def _do_bullets(self, curr_slide: ImpressPage[ImpressDoc]) -> None:
        # second slide: bullets and image
        body = curr_slide.bullets_slide(title="What is an Algorithm?")

        # bullet levels are 0, 1, 2,...
        body.add_bullet(
            level=0,
            text="An algorithm is a finite set of unambiguous instructions for solving a problem.",
        )

        body.add_bullet(
            level=1,
            text="An algorithm is correct if on all legitimate inputs, it outputs the right answer in a finite amount of time",
        )

        body.add_bullet(level=0, text="Can be expressed as")
        body.add_bullet(level=1, text="pseudocode")
        body.add_bullet(level=0, text="flow charts")
        body.add_bullet(
            level=1,
            text="text in a natural language (e.g. English)",
        )
        body.add_bullet(level=1, text="computer code")
        # add the image in bottom right corner, and scaled if necessary
        im = curr_slide.draw_image_offset(
            fnm=self._fnm_img,
            xoffset=ImageOffset(0.6),
            yoffset=ImageOffset(0.5),
        )
        # move below the slide text
        im.move_to_bottom()

    def _button_shapes(self, curr_slide: ImpressPage[ImpressDoc]) -> None:
        # fourth slide: title and rectangle (button) for playing a video
        # and a rounded button back to start
        curr_slide.title_only_slide("Wildlife Video Via Button")

        # button in the center of the slide
        sz = curr_slide.get_size_mm()
        width = 80
        height = 40

        ellipse = curr_slide.draw_ellipse(
            x=round((sz.Width - width) / 2),
            y=round((sz.Height - height) / 2),
            width=width,
            height=height,
        )

        ellipse.add_text(msg="Start Video", font_size=30)
        ellipse.set_property(
            OnClick=ClickAction.DOCUMENT, Bookmark=FileIO.fnm_to_url(self._fnm_wildlife)
        )
        # set Animation
        ellipse.set_property(
            Effect=AnimationEffect.MOVE_FROM_LEFT, Speed=AnimationSpeed.FAST
        )

        # draw a rounded rectangle with text
        button = curr_slide.draw_rectangle(
            x=sz.Width - width - 4,
            y=sz.Height - height - 5,
            width=width,
            height=height,
        )
        button.add_text(msg="Click to go\nto slide 1")
        button.set_gradient_color(name=DrawingGradientKind.SUNSHINE)
        # clicking makes the presentation jump to first slide
        button.set_property(CornerRadius=300, OnClick=ClickAction.FIRSTPAGE)

    def _dispatch_shapes(self, doc: ImpressDoc) -> None:
        curr_slide = doc.add_slide()
        curr_slide.title_only_slide("Dispatched Shapes")

        doc.set_visible()
        Lo.delay(1_000)

        doc.goto_page(page=curr_slide.component)
        Lo.print(
            f"Viewing Slide number: {Draw.get_slide_number(Draw.get_viewed_page(doc.component))}"
        )

        # first row
        y = 38
        _ = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.BASIC_SHAPES_DIAMOND,
            x=20,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        _ = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.THREE_D_HALF_SPHERE,
            x=80,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.CALLOUT_SHAPES_CLOUD_CALLOUT,
            x=140,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape.set_bitmap_color(name=DrawingBitmapKind.LITTLE_CLOUDS)

        dispatch_shape = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.FLOW_CHART_SHAPES_FLOWCHART_CARD,
            x=200,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape.set_hatch_color(name=DrawingHatchingKind.BLUE_NEG_45_DEGREES)
        # convert blue to black manually
        dispatch_hatch = cast(Hatch, dispatch_shape.get_property("FillHatch"))
        dispatch_hatch.Color = CommonColor.BLACK
        dispatch_shape.set_property(
            LineColor=CommonColor.BLACK, FillHatch=dispatch_hatch
        )
        # Props.show_obj_props("Hatch Shape", dispatch_shape)

        # second row
        y = 100
        dispatch_shape = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.STAR_SHAPES_STAR_12,
            x=20,
            y=y,
            width=40,
            height=40,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape.set_gradient_color(name=DrawingGradientKind.SUNSHINE)
        dispatch_shape.set_property(LineStyle=LineStyle.NONE)

        dispatch_shape = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.SYMBOL_SHAPES_HEART,
            x=80,
            y=y,
            width=40,
            height=40,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape.set_property(FillColor=CommonColor.RED)

        _ = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.ARROW_SHAPES_LEFT_RIGHT_ARROW,
            x=140,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape = curr_slide.add_dispatch_shape(
            shape_dispatch=ShapeDispatchKind.THREE_D_CYRAMID,
            x=200,
            y=y - 20,
            width=50,
            height=50,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dispatch_shape.set_bitmap_color(name=DrawingBitmapKind.STONE)

        Draw.show_shapes_info(curr_slide.component)
