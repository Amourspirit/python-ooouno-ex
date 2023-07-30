<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/185750033-5e0bc769-490f-4447-82fe-9badfa4ac208.svg" width="400" alt="replace"/>
</p>


# Text Replace

Example of Search and replace.

Search for the first occurrence of some words and/or replace some English spelled words with US spelled versions.
Displays a message box asking if you want to save document.
Optionally Saves the changed document in "replaced.doc" of working folder.

## See

See Also:

- [Text Search and Replace]
- [Finding the First Matching Phrase]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

Displays a message box asking if you want to save document.

Display a message box asking if you want to close document.

### Dev Container

From this folder.

```sh
python -m start --file "./data/bigStory.doc"
```

### Cross Platform

From this folder.

```sh
python -m start --file "./data/bigStory.doc"
```

### Linux/Mac

From project root folder (for default document).

```sh
python ./ex/auto/writer/odev_text_replace/start.py
```

### Windows

From project root folder (for default document).

```ps
python .\ex\auto\writer\odev_text_replace\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_text_replace'
```

This will copy the `odev_text_replace` example to the examples folder.

In the terminal run:

```bash
cd odev_text_replace
python -m start --file "./data/bigStory.doc"
```

[Text Search and Replace]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html
[Finding the First Matching Phrase]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html#finding-the-first-matching-phrase
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/

