# Make Table

Read tabbed text from an input file of Bond movies (bondMovies.txt)
and store as a blue table in "table.odt".

## See

See Also:

- [Text Content Other than Strings]
- [Adding a Text Table to a Document]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From project root folder.

```shell
python -m main auto -p "ex/auto/writer/odev_make_table/start.py"
```

### Linux

Run from current example folder.

```shell
python start.py
```

## Source (bondMovies.txt) CSV Data

```csv

// http://en.wikipedia.org/wiki/James_Bond#Ian_Fleming_novels

Title Year Actor Director

Dr. No 1962 Sean Connery Terence Young
From Russia with Love 1963 Sean Connery Terence Young
Goldfinger 1964 Sean Connery Guy Hamilton
Thunderball 1965 Sean Connery Terence Young
You Only Live Twice 1967 Sean Connery Lewis Gilbert
On Her Majesty's Secret Service 1969 George Lazenby Peter R. Hunt
Diamonds Are Forever 1971 Sean Connery Guy Hamilton
Live and Let Die 1973 Roger Moore Guy Hamilton
The Man with the Golden Gun 1974 Roger Moore Guy Hamilton
The Spy Who Loved Me 1977 Roger Moore Lewis Gilbert
Moonraker 1979 Roger Moore Lewis Gilbert
For Your Eyes Only 1981 Roger Moore John Glen
Octopussy 1983 Roger Moore John Glen
A View to a Kill 1985 Roger Moore John Glen
The Living Daylights 1987 Timothy Dalton John Glen
Licence to Kill 1989 Timothy Dalton John Glen
GoldenEye 1995 Pierce Brosnan Martin Campbell
Tomorrow Never Dies 1997 Pierce Brosnan Roger Spottiswoode
The World Is Not Enough 1999 Pierce Brosnan Michael Apted
Die Another Day 2002 Pierce Brosnan Lee Tamahori
Casino Royale 2006 Daniel Craig Martin Campbell
Quantum of Solace 2008 Daniel Craig Marc Forster
Skyfall 2012 Daniel Craig Sam Mendes
Spectre 2015 Daniel Craig Sam Mendes
```

## Screen shot of document

![image](https://user-images.githubusercontent.com/4193389/185208883-2a11e357-dde0-403a-ac08-b5696d51d5a9.png)

[Text Content Other than Strings]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter07.html
[Adding a Text Table to a Document]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter07.html#adding-a-text-table-to-a-document
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
