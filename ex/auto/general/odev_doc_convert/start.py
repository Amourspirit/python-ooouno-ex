#!/usr/bin/env python
# coding: utf-8
from __future__ import annotations
import argparse
from pathlib import Path

from ooodev.utils.lo import Lo
from ooodev.utils.info import Info
from ooodev.utils.file_io import FileIO


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument(
        "-e",
        "--ext",
        help="Extension of the converted file",
        action="store",
        dest="file_ext",
        required=True,
    )


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    # get extension from command line input
    ext = str(args.file_ext)

    # Using Lo.Loader context manager load Office and connect via socket.
    # Context manager takes care of terminating instance when job is done.
    # Note the use of the headless flag. Not using GUI for conversion.
    with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
        # get the absolute path of input file
        p_fnm = FileIO.get_absolute_path(args.file_path)

        name = Info.get_name(p_fnm)  # get name part of file without ext
        if not ext.startswith("."):
            # just in case user did not include . in --ext value
            ext = "." + ext

        p_save = Path(p_fnm.parent, f"{name}{ext}")  # new file, same as old file but different ext

        doc = Lo.open_doc(fnm=p_fnm, loader=loader)
        Lo.save_doc(doc=doc, fnm=p_save)
        Lo.close_doc(doc)

    print(f"All done! converted file: {p_save}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
