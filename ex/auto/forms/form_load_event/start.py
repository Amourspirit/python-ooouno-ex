from pathlib import Path

from load_listen import LoadListen
from ooodev.utils.file_io import FileIO


def main() -> int:
    db_fnm = Path("resources", "odb", "Example_Sport.odb")
    p = FileIO.get_absolute_path(db_fnm)
    if not p.exists():
        db_fnm = Path("../../../../resources", "odb", "Example_Sport.odb")
        p = FileIO.get_absolute_path(db_fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to Example_Sport.odb")
    _ = LoadListen(p)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
