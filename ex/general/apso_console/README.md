# APSO - Alternative Script Organizer for Python Example

This example requires that [APSO Extension](https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python) be installed into LibreOffice.

This example demonstrates how to load an **APSO** console on startup using a macro.

## Sample Document

See sample LibreOffice Writer document [apso_example.odt](apso_example.odt).

## Build

Build will compile the python scripts for this example into a single python script.

The following command will compile script as `apso_example.py` and embed it into `apso_example.odt`
The output is written into `build` folder in the projects root.

```sh
python -m main build -e --config "ex/general/apso_console/config.json" --embed-src "ex/general/apso_console/apso_example.odt"
```
