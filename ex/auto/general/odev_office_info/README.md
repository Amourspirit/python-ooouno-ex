# Office Info

This is a basic example that shows how to get info from Office via the command line.

This demo uses [OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/) (OooDev).

OooDev makes this demo possible with just a few lines of code.

See Also: [Examining Office](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part1/chapter03.html)

See [source code](./start.py)

## Example output

Print current user information

```sh
python -m start -u
```

```text
User Data Properties
  UserGroup: UserGroup
  apartment:
  c:
  encrypttoself: True
  facsimiletelephonenumber:
  fathersname:
  givenname: Amour
  homephone:
  initials: AS
  l:
  mail:
  o:
  position:
  postalcode:
  signingkey:
  sn: Spirit
  st:
  street:
  telephonenumber:
  title:

  Full Name: Amour Spirit
```

## Automate

### Dev Container

Run from current example folder.

```sh
python -m start
```

### Cross Platform

Run from current example folder.

```sh
python -m start
```

or for other commands

```sh
python start.py -h
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/general/odev_office_info/start.py
```

or for other commands

```sh
python ./ex/auto/general/odev_office_info/start.py -h
```

### Windows

From project root folder.

```ps
python .\ex\auto\general\odev_office_info\start.py
```

or for other commands

```ps
python .\ex\auto\general\odev_office_info\start.py -h
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/general/odev_office_info'
```

This will copy the `odev_office_info` example to the examples folder.

In the terminal run:

```bash
cd odev_office_info
python -m start
```

![command line example](https://user-images.githubusercontent.com/4193389/179056343-deafd3b5-c16e-45fa-9e2d-c95a0dc6b71e.gif)
