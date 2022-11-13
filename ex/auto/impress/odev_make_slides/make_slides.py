from __future__ import annotations
from typing import cast

import uno
from com.sun.star.lang import XComponent
from com.sun.star.drawing import XDrawPage

from ooodev.office.draw import (
    Draw,
    DrawingBitmapKind,
    DrawingHatchingKind,
    ShapeDispatchKind,
    DrawingGradientKind,
    ImageOffset,
)
from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
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
    def __init__(self, fnm_wildlife: PathOrStr, fnm_clock: PathOrStr, fnm_img: PathOrStr) -> None:
        _ = FileIO.is_exist_file(fnm=fnm_wildlife, raise_err=True)
        _ = FileIO.is_exist_file(fnm=fnm_clock, raise_err=True)
        _ = FileIO.is_exist_file(fnm=fnm_img, raise_err=True)
        self._fnm_wildlife = FileIO.get_absolute_path(fnm_wildlife)
        self._fnm_clock = FileIO.get_absolute_path(fnm_clock)
        self._fnm_img = FileIO.get_absolute_path(fnm_img)

    def main(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        try:
            doc = Draw.create_impress_doc(loader)
            curr_slide = Draw.get_slide(doc=doc, idx=0)

            GUI.set_visible(is_visible=True, odoc=doc)
            Lo.delay(1_000)  # delay to make sure zoom takes
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            Draw.title_slide(slide=curr_slide, title="Python-Generated Slides", sub_title="Using LibreOffice")

            # second slide
            curr_slide = Draw.add_slide(doc)
            self._do_bullets(curr_slide=curr_slide)

            # third slide: title and video
            curr_slide = Draw.add_slide(doc)
            Draw.title_only_slide(slide=curr_slide, header="Clock Video")
            Draw.draw_media(slide=curr_slide, fnm=self._fnm_clock, x=20, y=70, width=50, height=50)

            # fourth slide
            curr_slide = Draw.add_slide(doc)
            self._button_shapes(curr_slide=curr_slide)

            # fifth slide
            if DrawDispatcher:
                # windows only
                # a bit slow due to gui interaction but a good demo
                self._dispatch_shapes(doc)

            Lo.print(f"Total no. of slides: {Draw.get_slides_count(doc)}")

            Lo.delay(2000)
            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                Lo.close_doc(doc=doc, deliver_ownership=True)
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise

    def _do_bullets(self, curr_slide: XDrawPage) -> None:
        # second slide: bullets and image
        body = Draw.bullets_slide(slide=curr_slide, title="What is an Algorithm?")

        # bullet levels are 0, 1, 2,...
        Draw.add_bullet(
            bulls_txt=body,
            level=0,
            text="An algorithm is a finite set of unambiguous instructions for solving a problem.",
        )

        Draw.add_bullet(
            bulls_txt=body,
            level=1,
            text="An algorithm is correct if on all legitimate inputs, it outputs the right answer in a finite amount of time",
        )

        Draw.add_bullet(bulls_txt=body, level=0, text="Can be expressed as")
        Draw.add_bullet(bulls_txt=body, level=1, text="pseudocode")
        Draw.add_bullet(bulls_txt=body, level=0, text="flow charts")
        Draw.add_bullet(bulls_txt=body, level=1, text="text in a natural language (e.g. English)")
        Draw.add_bullet(bulls_txt=body, level=1, text="computer code")
        # add the image in bottom right corner, and scaled if necessary
        im = Draw.draw_image_offset(
            slide=curr_slide, fnm=self._fnm_img, xoffset=ImageOffset(0.6), yoffset=ImageOffset(0.5)
        )
        # move below the slide text
        Draw.move_to_bottom(slide=curr_slide, shape=im)

    def _button_shapes(self, curr_slide: XDrawPage) -> None:
        # fourth slide: title and rectangle (button) for playing a video
        # and a rounded button back to start
        Draw.title_only_slide(slide=curr_slide, header="Wildlife Video Via Button")

        # button in the center of the slide
        sz = Draw.get_slide_size(curr_slide)
        width = 80
        height = 40

        ellipse = Draw.draw_ellipse(
            slide=curr_slide,
            x=round((sz.Width - width) / 2),
            y=round((sz.Height - height) / 2),
            width=width,
            height=height,
        )

        Draw.add_text(shape=ellipse, msg="Start Video", font_size=30)
        Props.set(ellipse, OnClick=ClickAction.DOCUMENT, Bookmark=FileIO.fnm_to_url(self._fnm_wildlife))
        # set Animtion
        Props.set(
            ellipse,
            Effect=AnimationEffect.MOVE_FROM_LEFT,
            Speed=AnimationSpeed.FAST
        )

        # draw a rounded rectangle with text
        button = Draw.draw_rectangle(
            slide=curr_slide, x=sz.Width - width - 4, y=sz.Height - height - 5, width=width, height=height
        )
        Draw.add_text(shape=button, msg="Click to go\nto slide 1")
        Draw.set_gradient_color(shape=button, name=DrawingGradientKind.SUNSHINE)
        # clicking makes the presentation jump to first slide
        Props.set(button, CornerRadius=300, OnClick=ClickAction.FIRSTPAGE)

    def _dispatch_shapes(self, doc: XComponent) -> None:
        curr_slide = Draw.add_slide(doc)
        Draw.title_only_slide(slide=curr_slide, header="Dispatched Shapes")

        GUI.set_visible(is_visible=True, odoc=doc)
        Lo.delay(1_000)

        Draw.goto_page(doc=doc, page=curr_slide)
        Lo.print(f"Viewing Slide number: {Draw.get_slide_number(Draw.get_viewed_page(doc))}")

        # first row
        y = 38
        _ = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.BASIC_SHAPES_DIAMOND,
            x=20,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        _ = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.THREE_D_HALF_SPHERE,
            x=80,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dshape = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.CALLOUT_SHAPES_CLOUD_CALLOUT,
            x=140,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        Draw.set_bitmap_color(shape=dshape, name=DrawingBitmapKind.LITTLE_CLOUDS)

        dshape = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.FLOW_CHART_SHAPES_FLOWCHART_CARD,
            x=200,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        Draw.set_hatch_color(shape=dshape, name=DrawingHatchingKind.BLUE_NEG_45_DEGREES)
        # convert blue to black manually
        dhatch = cast(Hatch, Props.get(dshape, "FillHatch"))
        dhatch.Color = CommonColor.BLACK
        Props.set(dshape, LineColor=CommonColor.BLACK, FillHatch=dhatch)
        # Props.show_obj_props("Hatch Shape", dshape)

        # second row
        y = 100
        dshape = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.STAR_SHAPES_STAR_12,
            x=20,
            y=y,
            width=40,
            height=40,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        Draw.set_gradient_color(shape=dshape, name=DrawingGradientKind.SUNSHINE)
        Props.set(dshape, LineStyle=LineStyle.NONE)

        dshape = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.SYMBOL_SHAPES_HEART,
            x=80,
            y=y,
            width=40,
            height=40,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        Props.set(dshape, FillColor=CommonColor.RED)

        _ = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.ARROW_SHAPES_LEFT_RIGHT_ARROW,
            x=140,
            y=y,
            width=50,
            height=30,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        dshape = Draw.add_dispatch_shape(
            slide=curr_slide,
            shape_dispatch=ShapeDispatchKind.THREE_D_CYRAMID,
            x=200,
            y=y - 20,
            width=50,
            height=50,
            fn=DrawDispatcher.create_dispatch_shape,
        )
        Draw.set_bitmap_color(shape=dshape, name=DrawingBitmapKind.STONE)

        Draw.show_shapes_info(curr_slide)
