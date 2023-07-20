from __future__ import annotations
from pathlib import Path
from anim_bicycle import AnimBicycle

# region main()
def main() -> int:
    fnm = Path(__file__).parent / "image"/ "Bicycle-Blue.png"
    show = AnimBicycle(fnm_bike=fnm)
    show.animate()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
