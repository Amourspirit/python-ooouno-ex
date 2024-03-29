# Calc Add Range of Data Example

This example demonstrates how to add a range of data to a spreadsheet using a macro and [OOO Development Tools].

There is also an automation command line version of this example.
See [Add Range Automation](../../auto/calc/odev_add_range_data)

## Sample Document

See [odev_add_range.ods](odev_add_range.ods).

## Build

For automatic build run the following command from this folder.

```sh
make build
```

The following instructions are for manual build.

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `script.py` and embed it into `odev_add_range.ods`
The output is written into `build` folder in the projects root.

```sh
oooscript compile --embed --config "ex/calc/odev_add_range_data/config.json" --embed-doc "ex/calc/odev_add_range_data/odev_add_range.ods" -build-dir "build/add_range_data"
```

## Run Directly

To start LibreOffice and display a message box run the following command from this folder.

```sh
make run
```

![calc_range_macro](https://user-images.githubusercontent.com/4193389/173204999-924f12f6-59df-4bfe-8c2c-bee4cc5b9d6b.gif)

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
