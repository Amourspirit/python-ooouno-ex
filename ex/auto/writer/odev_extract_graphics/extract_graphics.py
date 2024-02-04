from __future__ import annotations
import tempfile
from pathlib import Path
from typing import cast

import uno
from com.sun.star.text import XTextShapesSupplier

from ooodev.office.draw import Draw
from ooodev.write import WriteDoc
from ooodev.loader import Lo
from ooodev.utils.file_io import FileIO
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.awt.size import Size


class ExtractGraphics:
    def __init__(self, fnm: PathOrStr, out_dir: PathOrStr = "") -> None:
        _ = FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)
        if out_dir:
            _ = FileIO.is_exist_dir(out_dir, True)
            self._out_dir = FileIO.get_absolute_path(out_dir)
        else:
            self._out_dir = Path(tempfile.mkdtemp())

    def main(self) -> None:
        with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
            text_doc = WriteDoc.open_doc(fnm=self._fnm, loader=loader)

            pics = text_doc.get_text_graphics()
            if not pics:
                text_doc.close_doc()
                return

            print()
            print(f"No. of text graphics: {len(pics)}")

            for i, pic in enumerate(pics):
                img_file = self._out_dir / f"graphics{i}.png"
                ImagesLo.save_graphic(pic=pic, fnm=img_file)
                sz = cast(Size, Props.get(pic, "SizePixel"))
                print(f"Image size in pixels: {sz.Width} X {sz.Height}")

            print()

            # this supplier is not created; Lo.qi() returns None
            shape_supp = text_doc.qi(XTextShapesSupplier)
            if shape_supp is None:
                print("Could not obtain text shapes supplier")
            else:
                print(f"No. of text shapes: {shape_supp.getShapes().getCount()}")

            # report on shapes in the doc
            shapes = text_doc.draw_page.get_shapes()
            if shapes:
                print()
                print(f"No. of draw shapes: {len(shapes)}")

                for shape in shapes:
                    Draw.report_pos_size(shape.component)
                print()

            text_doc.close_doc()
