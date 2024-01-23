from __future__ import annotations
import uno
from bezier_builder import BezierBuilder


def main() -> int:
    builder = BezierBuilder()
    builder.show()
    return 0


if __name__ == "__main__":
    SystemExit(main())
