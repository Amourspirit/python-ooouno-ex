from __future__ import annotations
import tempfile
from pathlib import Path

import uno
from ooodev.draw import ImpressDoc
from ooodev.utils.file_io import FileIO
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.lo import Lo
from ooodev.utils.type_var import PathOrStr
from ooodev.draw.filter.export_jpg import ExportJpg
from ooodev.draw.filter.export_png import ExportPng


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
        self._img_fmt = img_fmt.strip()
        self._resolution = resolution

    def main(self) -> None:
        # connect headless. will not need to see slides
        with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:
            doc = ImpressDoc.open_doc(fnm=self._fnm, loader=loader)

            slide = doc.slides[self._idx]

            names = ImagesLo.get_mime_types()
            Lo.print("Known GraphicExportFilter mime types:")
            for name in names:
                print(f"  {name}")

            out_fnm = self._out_dir / f"{self._fnm.stem}{self._idx}.{self._img_fmt}"
            print(f'Saving page {self._idx} to "{out_fnm}"')
            mime = ImagesLo.change_to_mime(self._img_fmt)

            # optionally set filter data to change resolution and have finer control.
            dt = None
            if mime == "image/jpeg":
                width, height = ImagesLo.get_dpi_width_height(
                    width=slide.component.Width,
                    height=slide.component.Height,
                    resolution=self._resolution,
                )
                dt = ExportJpg(
                    color_mode=True,
                    pixel_width=width,
                    pixel_height=height,
                    quality=80,
                    logical_width=width,
                    logical_height=height,
                ).to_filter_dict()
            if mime == "image/png":
                width, height = ImagesLo.get_dpi_width_height(
                    width=slide.component.Width,
                    height=slide.component.Height,
                    resolution=self._resolution,
                )
                # note: if `translucent=True` then page image are not exported,
                # also page margins are not exported.
                dt = ExportPng(
                    pixel_width=width,
                    pixel_height=height,
                    logical_width=width,
                    logical_height=height,
                    compression=6,
                    translucent=False,
                    interlaced=False,
                ).to_filter_dict()

            slide.save_page(fnm=out_fnm, mime_type=mime, filter_data=dt)
            doc.close_doc()
