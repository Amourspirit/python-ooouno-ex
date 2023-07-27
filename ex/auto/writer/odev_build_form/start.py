from __future__ import annotations
from pathlib import Path

from build_form import BuildForm


def main() -> int:
    db_fnm = Path(__file__).parent/ "data"  /"liang.odb"
    _ = BuildForm(db_fnm)

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
