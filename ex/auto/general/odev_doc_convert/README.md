# Write Convert Document Format

This is a basic example that opens a file and saves it as a new file type.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (ODEV).

ODEV makes this demo possible with just a few lines of code.

See [source code](./start.py)

## Automate

The following example command runs automation that reads the LICENSE file in this projects
root folder and writes out to LICENSE.pdf in the same folder.

```sh
python -m main auto --process 'ex/auto/general/odev_doc_convert/start.py -e "pdf" -f "LICENSE"'
```

![convert text to pdf](https://user-images.githubusercontent.com/4193389/178155989-1ec6e63a-ace3-4c60-8645-729245235d19.gif)
