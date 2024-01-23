from __future__ import annotations
import argparse
import sys

from slide_2_image import Slide2Image


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument(
        "-i",
        "--idx",
        help="Optional index of slide to convert to image. Default: %(default)i",
        action="store",
        dest="idx",
        type=int,
        default=0,
    )
    parser.add_argument(
        "-o",
        "--out_fmt",
        help="Extension of the converted file. Default: %(default)s",
        action="store",
        dest="output_format",
        default="jpg",
        choices=["jpg", "png"],
    )
    parser.add_argument(
        "-d",
        "--output_dir",
        help="Optional output Directory. Defaults to temporary dir sub folder.",
        action="store",
        dest="out_dir",
        default="",
    )

    parser.add_argument(
        "-r",
        "--resolution",
        help="Optional image output resolution in DPI. Defaults to 96.",
        action="store",
        dest="resolution",
        default=96,
        type=int,
    )


def _main() -> int:
    from pathlib import Path

    tmp = Path.cwd() / "tmp"
    tmp.mkdir(parents=True, exist_ok=True)

    if len(sys.argv) == 1:
        sys.argv.append("--file")
        sys.argv.append(str(Path(__file__).parent / "data/algs.ppt"))
        sys.argv.append("--idx")
        sys.argv.append("0")
        sys.argv.append("--out_fmt")
        sys.argv.append("jpg")
        sys.argv.append("--output_dir")
        sys.argv.append(str(tmp))
        sys.argv.append("--resolution")
        sys.argv.append("200")

    return main()


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    if len(sys.argv) == 1:
        parser.print_help()
        return 0
    args = parser.parse_args()
    sl = Slide2Image(
        fnm=args.file_path,
        idx=args.idx,
        img_fmt=args.output_format,
        out_dir=args.out_dir,
        resolution=args.resolution,
    )
    sl.main()

    return 0


if __name__ == "__main__":
    # switch to SystemExit(_main()) for debugging
    SystemExit(main())
