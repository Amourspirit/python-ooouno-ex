# Goal Seek

<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/205722662-32e8c4d1-678d-4906-812a-f207c724ddc0.png" width="575" height="345">
</p>

Examples of Goal Seek in a spreadsheet.

This demo uses This demo uses [OOO Development Tools] (ODEV).

See Also:

- [OOO Development Tools - Chapter 27. Functions and Data Analysis](https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part4/chapter27.html)

## Automate


### Cross Platform

From this folder.

```sh
python -m start
```

### Linux/Mac

```sh
python ./ex/auto/calc/odev_goal_seek/start.py
```

### Windows

```ps
python .\ex\auto\calc\odev_goal_seek\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/calc/odev_goal_seek'
```

This will copy the `odev_goal_seek` example to the examples folder.

In the terminal run:

```bash
cd odev_goal_seek
python -m start
```


## Output

```text
x == 16.0

'Divergence error: 1.7976931348623157e+308'

x == 1.0000000000000053 when x+1 == 2

x == 200000.0 when x*1.0*0.075 == 15000

x == -1.7692923428381226 when formula == 0
```

[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
