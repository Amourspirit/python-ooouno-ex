# Replace All

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/205417670-12dca447-60a7-4477-8853-c62cde6192e8.png" width="369" height="203">
</p>

Demonstrates Searching Iteratively; Searching For All Matches; Replacing All Matches in a spreadsheet.

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 26. Search and Replace](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter26.html)

## command line options

- `-a` Optional: Set search to use search all over search iter.
- `-r <word>` Optional: replacement word.
- `-o <out file>` Optional: file to save as
- search terms, one or more words to search.

### Example

Starts LibreOffice Calc, replaces all instance of `cat` and `hog` with `Bird` and saves output to `tmp/replaced_all.ods`

```sh
python -m start -a -o "tmp/replaced_all.ods" -r Bird cat hog
```

Starts

## Automate

A message box is display once the document has been processed asking if you want to close the document.

### Cross Platform

From this folder.

```sh
python -m start -h
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_replace_all/start.py -h
```

### Windows

```ps
python .\ex\auto\calc\odev_replace_all\start.py -h
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_replace_all'
```

This will copy the `odev_replace_all` example to the examples folder.

In the terminal run:

```bash
cd odev_replace_all
python -m start
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
