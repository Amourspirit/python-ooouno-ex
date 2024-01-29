from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

import uno

from ooodev.dialog.msgbox import (
    MsgBox,
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import CalcDoc, ZoomKind, CalcNamedEvent
from ooodev.utils.lo import Lo
from pathlib import Path

if TYPE_CHECKING:
    # Any imports in this block are only needed for type checking at design time and are not available at runtime.
    from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
    from ooodev.events.args.event_args_export import EventArgsExport
    from ooodev.calc.filter.export_jpg import ExportJpgT
    from ooodev.calc.filter.export_png import ExportPngT


def export_image(fnm: Path, resolution: int = 96) -> None:
    """
    Exports a range of cells as an image.

    Args:
        fnm (Path): The path to export image as (*.jpg or *.png).
        resolution (int, optional): Resolution to output image as. Defaults to 96.
    """
    data_file = Path(__file__).parent / "data" / "data.ods"

    loader = Lo.load_office(Lo.ConnectSocket())

    # region event handlers
    def on_exporting_png(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
        args.event_data["translucent"] = False
        args.event_data["compression"] = 8  # 0-9

    def on_exported_png(source: Any, args: EventArgsExport[ExportPngT]) -> None:
        print(f'Png URL: {args.get("url")}')

    def on_exporting_jpg(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
        args.event_data["quality"] = 80  # 0-100
        # when color_mode False image is exported as grayscale.
        args.event_data["color_mode"] = False

    def on_exported_jpg(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
        print(f'Jpg URL: {args.get("url")}')

    # endregion event handlers

    try:
        doc = CalcDoc.open_readonly_doc(fnm=data_file, loader=loader, visible=True)

        # delay before dispatching zoom
        Lo.delay(500)
        doc.zoom(ZoomKind.ZOOM_100_PERCENT)

        sheet = doc.get_active_sheet()

        # get a range of cells from the sheet
        rng = sheet.get_range(range_name="A1:N4")

        # Register event handlers so we can have a little more fine control over the export.
        # It is not required to register event handlers to export an image.
        rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_JPG, on_exporting_jpg)
        rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_JPG, on_exported_jpg)
        rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_PNG, on_exporting_png)
        rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_PNG, on_exported_png)

        # export range as image, in this case a png.
        # jpg format is also supported, eg: data.jpg
        rng.export_as_image(fnm=fnm, resolution=resolution)

        # see Also:
        # rng.export_jpg()
        # rng.export_png()

        # quick check to see if file exists
        assert fnm.exists()
        print(f"Image exported to {fnm}")

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


def main() -> int:
    # create a path to store file
    tmp = Path.cwd() / "tmp"
    tmp.mkdir(exist_ok=True)
    # change to jpg to export as jpg
    file = tmp / "data.png"
    # file = tmp / "data.jpg"
    export_image(file, 200)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
