from pathlib import Path

from build_form import BuildForm
from ooodev.utils.file_io import FileIO


def main() -> int:
    db_fnm = Path("resources", "odb", "liang.odb")
    p = FileIO.get_absolute_path(db_fnm)
    if not p.exists():
        db_fnm = Path("../../../../resources", "odb", "liang.odb")
        p = FileIO.get_absolute_path(db_fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to liang.odb")
    _ = BuildForm(p)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
