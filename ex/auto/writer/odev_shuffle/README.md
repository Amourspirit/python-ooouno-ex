# Shuffle Words

Example uses [OOO Development Tools].

Each word in the input file is mid-shuffled.
This causes the middle letters of the word to be rearranged, but not the first
and last letters. Words of <= 3 characters are unaffected.
The words are highlighted as they are shuffled.

Displays a message box asking if you want to save document.
If yes a file "shuffled.odt" is saved in the working folder.

## See

See Also:

- [Text API Overview]
- [Inserting/Changing Text in a Document]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

Displays a message box asking if you want to save document.

Display a message box asking if you want to close document.

### Dev Container

From this folder.

```shell
python -m start --file "./data/cicero_dummy.odt"
```

### Cross Platform

From this folder.

```shell
python -m start --file "./data/cicero_dummy.odt"
```

### Linux/Mac

From project root folder (for default file).

```sh
python ./ex/auto/writer/odev_shuffle/start.py
```

### Windows

From project root folder (for default file).

```ps
python .\ex\auto\writer\odev_shuffle\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_shuffle'
```

This will copy the `odev_shuffle` example to the examples folder.

In the terminal run:

```bash
cd odev_shuffle
python -m start
```

![shuffle text](https://user-images.githubusercontent.com/4193389/184251513-a8c96a5d-85b0-42ff-a891-ee5762e46a24.gif)

[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html

[Inserting/Changing Text in a Document]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html#inserting-changing-text-in-a-document
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
