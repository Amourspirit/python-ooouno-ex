from __future__ import annotations
import uno
from ooodev.loader import Lo
from ooodev.utils.gallery import Gallery, GalleryKind, SearchByKind, SearchMatchKind


class GalleryInfo:
    def main(self) -> None:
        with Lo.Loader(Lo.ConnectPipe(headless=True)):
            # list all the gallery themes (i.e. the sub-directories below gallery/)
            Gallery.report_galleries()
            print()

            # list all the items for the Sounds theme
            Gallery.report_gallery_items(GalleryKind.SOUNDS)
            print()

            # find an item that has "applause" as part of its name
            # in the Sounds theme
            itm = Gallery.find_gallery_obj(
                gallery_name=GalleryKind.SOUNDS,
                name="applause",
                search_match=SearchMatchKind.PARTIAL_IGNORE_CASE,
                search_kind=SearchByKind.FILE_NAME,
            )
            print()
            # print out the item's properties
            Gallery.report_gallery_item(itm)
