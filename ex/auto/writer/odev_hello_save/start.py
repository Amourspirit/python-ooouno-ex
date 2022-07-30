#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations

from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.gui import GUI


def main() -> int:

    # see: https://python-ooo-dev-tools.readthedocs.io/en/latest/src/utils/lo.html#ooodev.utils.lo.Lo.Loader
    with Lo.Loader(Lo.ConnectSocket()) as loader:

        doc = Write.create_doc(loader)

        GUI.set_visible(is_visible=True, odoc=doc)

        cursor = Write.get_cursor(doc)
        cursor.gotoEnd(False)  # make sure at end of doc before appending
        Write.append_para(cursor=cursor, text="Hello LibreOffice.\n")
        Lo.delay(1_000)  # Slow things down so user can see

        Write.append_para(cursor=cursor, text="How are you?")
        Lo.delay(2_000)
        Write.save_doc(text_doc=doc, fnm="hello.odt")
        Lo.close_doc(doc)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
