from __future__ import annotations
from pathlib import Path
from slide_chart import SlideChart


# region main()
def main() -> int:

    fnm = Path(__file__).parent / "data" / "chartsData.ods"

    sc = SlideChart(data_fnm=fnm)
    sc.main()
    return 0


# endregion main()

if __name__ == "__main__":
    SystemExit(main())
