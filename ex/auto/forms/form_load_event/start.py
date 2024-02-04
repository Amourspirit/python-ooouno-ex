from pathlib import Path

from load_listen import LoadListen


def main() -> int:
    db_fnm = Path(__file__).parent / "data" / "Example_Sport.odb"
    _ = LoadListen(db_fnm)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
