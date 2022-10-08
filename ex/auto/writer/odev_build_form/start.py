from pathlib import Path
import uno
from build_form import BuildForm
from ooodev.utils.file_io import FileIO
from ooodev.utils.lo import Lo

def main() -> int:
    db_fnm = Path("resources", "odb", "liang.odb")
    p = FileIO.get_absolute_path(db_fnm)
    if not p.exists():
        db_fnm = Path("../../../../resources", "odb", "liang.odb")
        p = FileIO.get_absolute_path(db_fnm)
    if not p.exists():
        raise FileNotFoundError("Unable to find path to liang.odb")
    frm = BuildForm(p)
    
    return 0

if __name__ == "__main__":
    raise SystemExit(main())