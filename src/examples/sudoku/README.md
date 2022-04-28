# LibreOffice Calc Sudoku

This is an example of Sudoku in **Calc**.

This example uses [types-scriptforge](https://pypi.org/project/types-scriptforge/) and [types-unopy](https://pypi.org/project/types-unopy/) give advantages of typing support and Intellisense support inside development environment.

To run the example you need to use LibreOffice >= `7.2`.

## Build

Build will compile the python scripts for this example into a single python script.


The following command will compile script as `calc-sudoku.py` and embed it into`calc-sudoku.ods`
The output is written into `build` folder in the projects root.

```sh
$ python -m main build -e --config 'src/examples/sudoku/config.json' --embed-src 'src/examples/sudoku/calc-sudoku.ods'
```

## Sample Document

See sample LibreOffice Calc document, [calc-sudoku.ods](calc-sudoku.ods)

![calc_sudoku](https://user-images.githubusercontent.com/4193389/165391098-883a7647-5fc8-47de-b028-4c2c98337abe.png)
