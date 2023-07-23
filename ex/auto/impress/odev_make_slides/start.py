from __future__ import annotations
from pathlib import Path
from make_slides import MakeSlides

# region main()
def main() -> int:
    data_dir = Path(__file__).parent / "data"
    fnm_img = data_dir / "skinner.png"
    fnm_clock = data_dir / "clock.avi"
    fnm_wildlife = data_dir / "wildlife.mp4"

    m_slides = MakeSlides(fnm_wildlife=fnm_wildlife, fnm_clock=fnm_clock, fnm_img=fnm_img)
    m_slides.main()
    return 0

# endregion main()

if __name__ == "__main__":
    SystemExit(main())
