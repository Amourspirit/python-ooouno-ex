# LibreOffice Calc Sudoku

This is an example of Sudoku in **Calc** using [ScriptForge].

This example uses [types-scriptforge](https://pypi.org/project/types-scriptforge/) and [types-unopy](https://pypi.org/project/types-unopy/) give advantages of typing support and Intellisense (autocomplete) support inside development environment.

To run the example you need to use LibreOffice >= `7.2`.

## Build

`make build` will compile the python scripts for this example into a single python script and
embed the script into the `calc-sudoku.ods` document in the `build/sudoku` folder.

From the current folder run:

```sh
make build
```

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html).

## Sample Document

See sample LibreOffice Calc document, [calc-sudoku.ods](calc-sudoku.ods)

![calc_sudoku](https://user-images.githubusercontent.com/4193389/165391098-883a7647-5fc8-47de-b028-4c2c98337abe.png)

[ScriptForge]: https://gitlab.com/LibreOfficiant/scriptforge
