from __future__ import annotations
from pathlib import Path
from text_chart import TextChart


# region main()
def main() -> int:

    fnm = Path(__file__).parent / "data" / "chartsData.ods"

    tc = TextChart(data_fnm=fnm)
    tc.main()
    return 0

# endregion main()

if __name__ == "__main__":
    SystemExit(main())
