<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/184450252-a05db707-4e64-496d-8d1b-0cb6a4792361.svg" width="400" alt="folder image"/>
</p>

# Docs Append

This example uses [OOO Development Tools] to show how to append other documents to an existing document.

Example uses `XDocumentInsertable.insertDocumentFromURL()`. A list of filenames is read from the command line; the first file is opened, and the other files appended to it by `append_text_files()`:

The automate line below opens `blank.odt`, appends `story.odt` and then `cicero_dummy.odt`.
The result is saved in a new document in the current working folder and is named `blank_APPENDED.odt`.

## See

See Also:

- [Text API Overview]
- [Appending Documents Together]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Dev Container

From current example folder.

```shell
python -m start --file "./data/blank.odt" "./data/story.odt" "./data/cicero_dummy.odt"
```

### Cross Platform

From current example folder.

```shell
python -m start --file "./data/blank.odt" "./data/story.odt" "./data/cicero_dummy.odt"
```

### Linux/Mac

From project root folder

```sh
python ./ex/auto/writer/odev_docs_append/start.py -f "ex/auto/writer/odev_docs_append/data/blank.odt" "ex/auto/writer/odev_docs_append/data/story.odt" "ex/auto/writer/odev_docs_append/data/cicero_dummy.odt"
```

### Windows

From project root folder

```ps
python .\ex\auto\writer\odev_docs_append\start.py -f "ex/auto/writer/odev_docs_append/data/blank.odt" "ex/auto/writer/odev_docs_append/data/story.odt" "rex/auto/writer/odev_docs_append/data/cicero_dummy.odt"
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_docs_append'
```

This will copy the `odev_docs_append` example to the examples folder.

In the terminal run:

```bash
cd odev_docs_append
python -m start
```


[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[Appending Documents Together]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html#appending-documents-together
