import sys
from ooodev.utils.file_io import FileIO
from make_slides import MakeSlides

# region maind()
def main() -> int:
    fnm_img = "resources/image/skinner.png"
    if not FileIO.is_exist_file(fnm_img):
        fnm_img = "../../../../resources/image/skinner.png"
        
    fnm_clock = "resources/video/clock.avi"
    if not FileIO.is_exist_file(fnm_clock):
        fnm_clock = "../../../../resources/video/clock.avi"


    fnm_wildlife = "resources/video/wildlife.mp4"
    if not FileIO.is_exist_file(fnm_wildlife):
        fnm_wildlife = "../../../../resources/video/wildlife.mp4"

    m_slides = MakeSlides(fnm_wildlife=fnm_wildlife, fnm_clock=fnm_clock, fnm_img=fnm_img)
    m_slides.main()
    return 0


# endregion maind()

if __name__ == "__main__":
    SystemExit(main())
