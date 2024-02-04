from __future__ import annotations
from typing import Any, TYPE_CHECKING
import tempfile
from pathlib import Path

import uno
from ooodev.draw import ImpressDoc, DrawNamedEvent
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.loader import Lo
from ooodev.utils.file_io import FileIO
from ooodev.utils.type_var import PathOrStr

if TYPE_CHECKING:
    from ooodev.draw.filter.export_png import ExportPngT
    from ooodev.draw.filter.export_jpg import ExportJpgT


class Slide2Image:
    def __init__(
        self,
        fnm: PathOrStr,
        idx: int,
        img_fmt: str,
        out_dir: PathOrStr = "",
        resolution: int = 96,
    ) -> None:
        _ = FileIO.is_exist_file(fnm, True)
        self._fnm = FileIO.get_absolute_path(fnm)
        if out_dir:
            _ = FileIO.is_exist_dir(out_dir, True)
            self._out_dir = FileIO.get_absolute_path(out_dir)
        else:
            self._out_dir = Path(tempfile.mkdtemp())
        if idx < 0:
            print("Index is less then zero. Using zero")
            idx = 0
        self._idx = idx
        self._img_fmt = img_fmt.strip().lower()
        self._resolution = resolution

    def main(self) -> None:
        def on_exporting_jpg(
            source: Any, args: CancelEventArgsExport[ExportJpgT]
        ) -> None:
            args.event_data["quality"] = 80

        def on_exported_jpg(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
            print("Exported jpg")

        def on_exporting_png(
            source: Any, args: CancelEventArgsExport[ExportPngT]
        ) -> None:
            args.event_data["compression"] = 7

        def on_exported_png(source: Any, args: EventArgsExport[ExportPngT]) -> None:
            print("Exported png")

        # connect headless. will not need to see slides
        with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:
            doc = ImpressDoc.open_doc(fnm=self._fnm, loader=loader)
            slide = doc.slides[self._idx]

            # Optionally subscribe to events. This allow more control over the export process.
            slide.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_JPG, on_exporting_jpg)
            slide.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_JPG, on_exported_jpg)
            slide.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_PNG, on_exporting_png)
            slide.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_PNG, on_exported_png)

            out_fnm = self._out_dir / f"{self._fnm.stem}{self._idx}.{self._img_fmt}"
            print(f'Saving page {self._idx} to "{out_fnm}"')
            if self._img_fmt == "png":
                slide.export_page_png(out_fnm, self._resolution)
            else:
                slide.export_page_jpg(out_fnm, self._resolution)
            doc.close_doc()
