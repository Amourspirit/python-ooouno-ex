# LibreOffice Conference 2021

This Example contains the examples used in the talk "Python scripts in LibreOffice Calc using the [ScriptForge] library" given by Rafael Lima during the LibreOffice Conference 2021.
Original on [github](https://github.com/rafaelhlima/LibOCon_2021_SFCalc).

This example uses [types-scriptforge]) and [types-unopy] give advantages of typing support and Intellisense (autocomplete) support inside development environment.

Adding type support to this example is the reason is has also been duplicated from the [original](https://github.com/rafaelhlima/LibOCon_2021_SFCalc).

To run the examples in the Python file you need to use LibreOffice >= `7.2`.

The [lib_o_con_2021.ods](./lib_o_con_2021.ods) is a working example.\
For the supporting files the res folder must be located in the same folder as [lib_o_con_2021.ods](./lib_o_con_2021.ods).

[LibreOffice Conference 2021 on YouTube](https://youtu.be/3xnO1prvgmk).

## Build

If not on windows you can run the following command to build the example from this folder.

```sh
make build
```

The following instructions are for manual build.

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `lib_o_con_2021.py` and embed it into `lib_o_con_2021.ods`
The output is written into `build/lib_o_con_2021` folder in the projects root.

```sh
oooscript compile --embed --config "ex/calc/lib_o_con_2021/config.json" --embed-doc "ex/calc/lib_o_con_2021/lib_o_con_2021.ods" --build-dir "build/lib_o_con_2021"
```

Also the `res` folder of this example needs copied to the `build/lib_o_con_2021` folder.


## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd macro
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/calc/lib_o_con_2021'
```

This will copy the `lib_o_con_2021` example to the macro folder.

## Screenshot

![lib_o_con_2021_ods](https://user-images.githubusercontent.com/4193389/163496918-1f0a171c-b939-4f18-b674-a9b4cd35fc5a.png)

[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
[types-scriptforge]: https://pypi.org/project/types-scriptforge/
[types-unopy]: https://pypi.org/project/types-unopy/
