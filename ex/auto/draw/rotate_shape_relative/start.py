from __future__ import annotations
from pathlib import Path
from rotate_shape import RotateShape


# region main()
def main() -> int:
    show = RotateShape()
    show.animate()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
