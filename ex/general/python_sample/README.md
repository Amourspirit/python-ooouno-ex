# Python Sample

<p align="center">
<img src="https://github.com/Amourspirit/python-ooouno-ex/assets/4193389/ad49fb93-badb-4d28-96bc-d90266d1fac2" width="650" height="470">
</p>

## Overview

This a simple python sample project.
The purpose of this sample is to demonstrate how to embed a simple python script into a LibreOffice document.

This sample uses the [oooscript] library to embed the python script into the document.

## Building

To build the sample, run the following command in a terminal of the current folder:

```bash
make build
```

The `make build` command invokes [oooscript].

This will embed `sample.py` into the `data/sample.odt` file and save it as `build/python_sample/python_sample.odt`.
The `build` is in the root of this project. It will be created if it does not exist.


## Notes

The `data/sample.odt` file contains an embedded dialog named `Dialog1`.
The  `sample.py` script contains the `doc_dialog` method that is used to open the dialog.

See also LibreOffice Help [Opening a Dialog with Python](https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_dialogs.html?&DbPAR=WRITER)

[oooscript]: https://oooscript.readthedocs.io/en/latest/