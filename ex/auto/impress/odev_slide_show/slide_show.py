from __future__ import annotations

import uno
from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.draw import (
    Draw,
    ImpressDoc,
    FadeEffect,
    AnimationSpeed,
    DrawingGradientKind,
    DrawingSlideShowKind,
)
from ooodev.utils.dispatch.draw_view_dispatch import DrawViewDispatch
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props

from ooo.dyn.presentation.animation_effect import AnimationEffect
from ooo.dyn.presentation.click_action import ClickAction


class SlideShow:
    def show(self) -> None:
        loader = Lo.load_office(Lo.ConnectPipe())

        # create Impress page or Draw slide
        try:
            doc = ImpressDoc(Draw.create_impress_doc(loader))

            while doc.get_slides_count() < 3:
                _ = doc.add_slide()

            # ---- The first page
            slide = doc.get_slide(idx=0)
            slide.set_transition(
                fade_effect=FadeEffect.FADE_FROM_RIGHT,
                speed=AnimationSpeed.FAST,
                change=DrawingSlideShowKind.AUTO_CHANGE,
                duration=1,
            )
            # draw a square at the top left of the page; and text
            sq1 = slide.draw_rectangle(x=10, y=10, width=50, height=50)
            sq1.set_property(Effect=AnimationEffect.WAVYLINE_FROM_BOTTOM)
            # square appears in 'wave' (pixels by pixels)
            _ = slide.draw_text(
                msg="Page 1", x=70, y=20, width=60, height=30, font_size=24
            )

            # ---- The second page
            slide = doc.get_slide(idx=1)
            slide.set_transition(
                fade_effect=FadeEffect.FADE_FROM_RIGHT,
                speed=AnimationSpeed.FAST,
                change=DrawingSlideShowKind.AUTO_CHANGE,
                duration=1,
            )
            # draw a circle at the bottom right of second page; and text
            circle1 = slide.draw_ellipse(x=212, y=150, width=50, height=50)
            # hide circle after drawing
            circle1.set_property(Effect=AnimationEffect.HIDE)

            _ = slide.draw_text(
                msg="Page 2",
                x=170,
                y=170,
                width=60,
                height=30,
                font_size=24,
            )

            name_slide = "page two"
            slide.set_name(name_slide)

            # ---- The third page
            slide = doc.get_slide(idx=2)
            slide.set_transition(
                fade_effect=FadeEffect.ROLL_FROM_LEFT,
                speed=AnimationSpeed.MEDIUM,
                change=DrawingSlideShowKind.AUTO_CHANGE,
                duration=2,
            )
            _ = slide.draw_text(
                msg="Page 3",
                x=120,
                y=75,
                width=60,
                height=30,
                font_size=24,
            )
            # draw a circle containing text
            circle2 = slide.draw_ellipse(x=10, y=8, width=50, height=50)
            circle2.add_text(msg="Click to go\nto Page1")
            circle2.set_gradient_color(name=DrawingGradientKind.MAHOGANY)

            # clicking makes the presentation jump to page one
            circle2.set_property(
                Effect=AnimationEffect.FADE_FROM_BOTTOM,
                OnClick=ClickAction.FIRSTPAGE,
            )

            # draw a square with text
            sq2 = slide.draw_rectangle(x=220, y=8, width=50, height=50)
            sq2.add_text(msg="Click to go\nto Page 2")
            sq2.set_gradient_color(name=DrawingGradientKind.MAHOGANY)

            # clicking makes the presentation jump to page two by using a bookmark
            sq2.set_property(
                Effect=AnimationEffect.FADE_FROM_BOTTOM,
                OnClick=ClickAction.BOOKMARK,
                Bookmark=name_slide,
            )

            # slideshow start() crashes if the doc is not visible
            doc.set_visible()
            show = doc.get_show()

            # a full-screen slide show
            Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION)
            Lo.delay(500)
            # show.start() starts slideshow but not necessarily in 100% full screen
            # show.start()

            Props.show_obj_props("Slide show", show)
            sc = doc.get_show_controller()

            # Draw.wait_last(sc, 2000)
            # Lo.dispatch_cmd(DrawViewDispatch.PRESENTATION_END)
            Draw.wait_ended(sc)
            Lo.delay(1_000)
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
