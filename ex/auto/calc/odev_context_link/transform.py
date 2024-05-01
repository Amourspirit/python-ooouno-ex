#!/usr/bin/env python
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import re
import contextlib
import uno
from ooo.dyn.util.url import URL
from ooodev.calc import CalcDoc, ZoomKind

from ooodev.loader import Lo
from ooodev.adapter.util.url_transformer_comp import URLTransformerComp


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from ooodev.events.args.event_args import EventArgs


def main() -> int:
    _ = Lo.load_office(Lo.ConnectSocket())
    try:
        doc = CalcDoc.create_doc(visible=True)
        tf = URLTransformerComp.from_lo()
        # surl = ".uno:ooodev.Freeline_Unfilled?Transparence:short=50&Color:string=COL_GRAY7&Width:short=500&IsSticky:bool=true&ShapeName:string=FreeformRedactionShape"
        surl = ".uno:ooodev.calc.menu.convert.url?cell=A1&rice=True"
        url = URL()
        url.Complete = surl
        b, result = tf.parse_strict(url)
        print(result)
        doc.close_doc()
        Lo.close_office()

    except Exception:
        Lo.close_office()
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
