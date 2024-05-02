# APSO - Alternative Script Organizer for Python Example

This example requires that [APSO Extension] be installed into LibreOffice.

This example demonstrates how to load an **APSO** console on startup using a macro.

## Sample Document

See sample LibreOffice Writer document [apso_example.odt](apso_example.odt).

## Build

For automatic build run the following command from this folder.

```sh
make build
```

The following instructions are for manual build.

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `apso_example.py` and embed it into `apso_example.odt`
The output is written into `build` folder in the projects root.

```sh
oooscript compile --embed --config "ex/general/apso_console/config.json" --embed-doc "ex/general/apso_console/apso_example.odt" -build-dir "build/apso_console"
```

See [Guide on embedding python macros in a LibreOffice Document](https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/embed_python.html).

[APSO Extension]: https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python
