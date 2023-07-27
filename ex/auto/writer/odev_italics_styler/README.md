<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/185763894-adb25e29-270f-4085-834b-502cf48c86fe.png" alt="Italics"/>
</p>

# Italics Styler

Example of Search and style Italics.

By adding `--word name color` instances to the command line determines which words are italicize and what color the italicize word shall be.

In this example:
Search for all occurrences of a `pleasure` (case-insensitive) and set their style to be in green italics.
Search for all occurrences of a `pain` (case-insensitive) and set their style to be in red italics.

Save changed  text to "italicized.doc" in current working folder.

## See

See Also:

- [Text Search and Replace]
- [Finding the First Matching Phrase]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

Displays a message box asking if you want to save document.

Display a message box asking if you want to close document.

### Cross Platform

From this folder.

```sh
python -m start --file "./data/cicero_dummy.odt" --word pleasure green --word pain red
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/writer/odev_italics_styler/start.py --file "ex/auto/writer/odev_italics_styler/data/cicero_dummy.odt" --word pleasure green --word pain red
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_italics_styler\start.py --file "ex/auto/writer/odev_italics_styler/data/cicero_dummy.odt" --word pleasure green --word pain red
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_italics_styler'
```

This will copy the `odev_italics_styler` example to the examples folder.

In the terminal run:

```bash
cd odev_italics_styler
python -m start
```

## Output

```text
Searching for all occurrences of 'pleasure'
No. of matches: 17
  - found: 'pleasure'
    - on page 1
    - starting at char position: 85
  - found: 'pleasure'
    - on page 1
    - starting at char position: 319
  - found: 'pleasure'
    - on page 1
    - starting at char position: 350
  - found: 'pleasure'
    - on page 1
    - starting at char position: 408
  - found: 'pleasure'
    - on page 1
    - starting at char position: 679
  - found: 'pleasure'
    - on page 1
    - starting at char position: 884
  - found: 'pleasure'
    - on page 1
    - starting at char position: 980
  - found: 'pleasure'
    - on page 1
    - starting at char position: 1118
  - found: 'pleasure'
    - on page 1
    - starting at char position: 1571
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2057
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2291
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2322
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2380
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2651
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2856
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2952
  - found: 'pleasure'
    - on page 1
    - starting at char position: 3089
Found 17 results for "pleasure"
Searching for all occurrences of 'pain'
No. of matches: 15
  - found: 'pain'
    - on page 1
    - starting at char position: 107
  - found: 'pain'
    - on page 1
    - starting at char position: 548
  - found: 'pain'
    - on page 1
    - starting at char position: 578
  - found: 'pain'
    - on page 1
    - starting at char position: 647
  - found: 'pain'
    - on page 1
    - starting at char position: 948
  - found: 'pain'
    - on page 1
    - starting at char position: 1193
  - found: 'pain'
    - on page 1
    - starting at char position: 1377
  - found: 'pain'
    - on page 1
    - starting at char position: 1608
  - found: 'pain'
    - on page 1
    - starting at char position: 2079
  - found: 'pain'
    - on page 1
    - starting at char position: 2520
  - found: 'pain'
    - on page 1
    - starting at char position: 2550
  - found: 'pain'
    - on page 1
    - starting at char position: 2619
  - found: 'pain'
    - on page 1
    - starting at char position: 2920
  - found: 'pain'
    - on page 1
    - starting at char position: 3164
  - found: 'pain'
    - on page 1
    - starting at char position: 3348
Found 15 results for "pain"
```

[Text Search and Replace]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html
[Finding the First Matching Phrase]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html#finding-the-first-matching-phrase
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
