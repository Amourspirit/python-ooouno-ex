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

### Cross Platform

From project root folder.

```shell
python -m main auto -p "ex/auto/writer/odev_docs_append/start.py --file resources/odt/blank.odt resources/odt/story.odt resources/odt/cicero_dummy.odt"
```

### Linux

Run from current example folder.

```shell
python start.py --file "../../../../resources/odt/blank.odt" "../../../../resources/odt/story.odt" "../../../../resources/odt/cicero_dummy.odt"
```

[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
[Appending Documents Together]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html#appending-documents-together
