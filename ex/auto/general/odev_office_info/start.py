#!/usr/bin/env python
# coding: utf-8
# region Imports
from __future__ import annotations
import argparse
import platform

from ooodev.utils.lo import Lo
from ooodev.utils.info import Info
from ooodev.utils.file_io import FileIO

# endregion Imports

# region args
def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-s",
        "--services",
        help="Show Services",
        action="store_true",
        dest="services",
        default=False,
    )
    parser.add_argument(
        "-w",
        "--write-services",
        help="Show Write Services",
        action="store_true",
        dest="services_write",
        default=False,
    )
    parser.add_argument(
        "-c",
        "--calc-services",
        help="Show Calc Services",
        action="store_true",
        dest="services_calc",
        default=False,
    )
    parser.add_argument(
        "-f",
        "--filters",
        help="Show Filters",
        action="store_true",
        dest="filters",
        default=False,
    )


# endregion args

# region Show


def show_filters() -> None:
    print()
    print(" File Filter Names for Office: ".center(50, "-"))
    filters = Info.get_filter_names()
    for filter in filters:
        print(f"  {filter}")
    print("-----------")
    print(f"No. of filters: {len(filter)}")


def show_services(title: str, srv_name: str | None = None) -> None:
    print()
    print(title.center(50, "-"))
    services = Info.get_service_names(srv_name)
    for service in services:
        print(f"  {service}")
    print("-----------")
    print(f"No. of services: {len(services)}")


# endregion Show

# region Main
def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    # read the current command line args
    args = parser.parse_args()

    with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:
        print(f"OS Platform: {platform.platform()}")
        print(f"OS Version: {platform.version()}")
        print(f"OS Release: {platform.release()}")
        print(f"OS Architecture: {platform.architecture()}")

        print(f"\nOffice Name: {Info.get_config('ooName')}")
        print(f"\nOffice version (long): {Info.get_config('ooSetupVersionAboutBox')}")
        print(f"Office version (short): {Info.get_config('ooSetupVersion')}")
        print(f"\nOffice language location: {Info.get_config('ooLocale')}")
        print(f"System language location: {Info.get_config('ooSetupSystemLocale')}")

        print(f"\nWorking Dir: {Info.get_paths('Work')}")
        addin_dir = Info.get_paths('Addin')
        print(f"\nAddin Dir: {addin_dir}")
        print(f"Addin Path: {FileIO.uri_to_path(addin_dir)}")
        print(f"\nOffice Dir: {Info.get_office_dir()}")
        print(f"\nFilters Dir: {Info.get_paths('Filter')}")
        print(f"\nTemplates Dirs: {Info.get_paths('Template')}")
        print(f"\nGallery Dir: {Info.get_paths('Gallery')}")
        # see: https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings

        if args.filters:
            show_filters()
        if args.services:
            show_services(" Services for Office: ")
            # over 900 printed
        if args.services_write:
            show_services(" Services for Write: ", Lo.Service.WRITER)
        if args.services_calc:
            show_services(" Services for Calc: ", Lo.Service.CALC)

        print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# endregion main
