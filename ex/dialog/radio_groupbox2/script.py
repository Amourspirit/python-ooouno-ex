from radio_group_box import RadioGroupBox

# import the WriteDoc class just to ensure that it is included in the full script compile.
from ooodev.write import WriteDoc


def show_radio_dialog(*args) -> None:
    rgb = RadioGroupBox()
    rgb.show()
