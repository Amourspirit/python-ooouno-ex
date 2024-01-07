from __future__ import annotations
import uno
from ooodev.format.inner.direct.write.frame.frame_type.position import RelVertOrient
from ooodev.write import Write, WriteDoc, ZoomKind
from ooodev.utils.lo import Lo
from ooodev.units import UnitMM100

from ooodev.format.writer.direct.obj.type import (
    Anchor,
    AnchorKind,
    Position,
    Horizontal,
    HoriOrient,
    Vertical,
    VertOrient,
    RelHoriOrient,
)
from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.format.writer.direct.obj.wrap import (
    Settings as WrapSettings,
    WrapTextMode,
    Options as WrapOptions,
)
from ooodev.format.draw.direct.area import (
    Gradient as DrawAreaGradient,
    PresetGradientKind,
)


def main() -> int:
    """Main Entry Point"""
    p_txt = (
        "To Sherlock Holmes she is always THE woman. I have seldom heard"
        " him mention her under any other name. In his eyes she eclipses"
        " and predominates the whole of her sex. It was not that he felt"
        " any emotion akin to love for Irene Adler. All emotions, and that"
        " one particularly, were abhorrent to his cold, precise but"
        " admirably balanced mind. He was, I take it, the most perfect"
        " reasoning and observing machine that the world has seen, but as a"
        " lover he would have placed himself in a false position."
        " He never spoke of the softer passions, save with a gibe and a sneer."
        " They were admirable things for the observer--excellent for drawing the"
        " veil from men's motives and actions. But for the trained reasoner"
        " to admit such intrusions into his own delicate and finely"
        " adjusted temperament was to introduce a distracting factor which"
        " might throw a doubt upon all his mental results."
    )
    loader = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = WriteDoc(Write.create_doc(loader))
        doc.set_visible()

        # delay so document is visible before dispatching zoom
        Lo.delay(500)
        doc.zoom(ZoomKind.ENTIRE_PAGE)

        cursor = doc.get_cursor()

        cursor.append(p_txt)
        tvc = doc.get_view_cursor()
        tvc.goto_end()
        current_pos = tvc.get_position()

        doc.draw_page

        radius = UnitMM100.from_mm(15)
        rect = doc.draw_page.draw_circle(1, 1, radius)
        rect.set_string("Round we go!")
        assert rect is not None

        sz = doc.get_page_size()  # in 100th of mm units
        center = sz.width // 2

        rect_horz = UnitMM100(center - radius.value)
        rect_vert = UnitMM100(current_pos.Y)
        anchor = Anchor(anchor=AnchorKind.AT_CHARACTER)  # anchor to text
        # get a gradient from a preset to apply to rect shape
        gradient = DrawAreaGradient.from_preset(preset=PresetGradientKind.GREEN_GRASS)

        # position the shape
        pos_style = Position(
            vertical=Vertical(
                position=VertOrient.FROM_TOP_OR_BOTTOM,
                rel=RelVertOrient.ENTIRE_PAGE_OR_ROW,
                amount=rect_vert,
            ),
            horizontal=Horizontal(
                position=HoriOrient.FROM_LEFT_OR_INSIDE,
                rel=RelHoriOrient.ENTIRE_PAGE,
                amount=rect_horz,
            ),
        )
        wrap_style = WrapSettings(mode=WrapTextMode.PARALLEL)  # wrap text around shape
        wrap_opt = WrapOptions(contour=True)  # contour text around shape

        # apply the styles to the rect shape
        rect.apply_styles(anchor, gradient, pos_style, wrap_style, wrap_opt)

        cursor.append_para(
            " Grit in a sensitive instrument, or a crack in one of his own high-power"
            " lenses, would not be more disturbing than a strong emotion in a"
            " nature such as his. And yet there was but one woman to him, and"
            " that woman was the late Irene Adler, of dubious and questionable memory."
        )
        # pause for 1 second
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
