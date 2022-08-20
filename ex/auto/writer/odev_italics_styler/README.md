<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/185751451-e1544108-f0c8-4e38-86c9-2b0f9b57a94e.png" alt="Italics"/>
</p>

# Italics Styler

Example of Search and style Italics.

Search for all occurrences of a `pleasure` (case-insensitive) and set their style to be in green italics.
Search for all occurrences of a `pain` (case-insensitive) and set their style to be in red italics.
Save changed  text to "italicized.doc" in current folder.

## See

See Also:

- [Text Search and Replace]
- [Finding the First Matching Phrase]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

Run from current example folder.

```shell
python start.py --show --file "../../../../resources/odt/cicero_dummy.odt"
```

## Output

```text
Searching for all occurrences of 'pleasure'
No. of matches: 17
  - found: 'pleasure'
    - on page 1
    - starting at char position: 85
  - found: 'pleasure'
    - on page 1
    - starting at char position: 319
  - found: 'pleasure'
    - on page 1
    - starting at char position: 350
  - found: 'pleasure'
    - on page 1
    - starting at char position: 408
  - found: 'pleasure'
    - on page 1
    - starting at char position: 679
  - found: 'pleasure'
    - on page 1
    - starting at char position: 884
  - found: 'pleasure'
    - on page 1
    - starting at char position: 980
  - found: 'pleasure'
    - on page 1
    - starting at char position: 1118
  - found: 'pleasure'
    - on page 1
    - starting at char position: 1571
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2057
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2291
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2322
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2380
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2651
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2856
  - found: 'pleasure'
    - on page 1
    - starting at char position: 2952
  - found: 'pleasure'
    - on page 1
    - starting at char position: 3089
Found 17 results for "pleasure"
Searching for all occurrences of 'pain'
No. of matches: 15
  - found: 'pain'
    - on page 1
    - starting at char position: 107
  - found: 'pain'
    - on page 1
    - starting at char position: 548
  - found: 'pain'
    - on page 1
    - starting at char position: 578
  - found: 'pain'
    - on page 1
    - starting at char position: 647
  - found: 'pain'
    - on page 1
    - starting at char position: 948
  - found: 'pain'
    - on page 1
    - starting at char position: 1193
  - found: 'pain'
    - on page 1
    - starting at char position: 1377
  - found: 'pain'
    - on page 1
    - starting at char position: 1608
  - found: 'pain'
    - on page 1
    - starting at char position: 2079
  - found: 'pain'
    - on page 1
    - starting at char position: 2520
  - found: 'pain'
    - on page 1
    - starting at char position: 2550
  - found: 'pain'
    - on page 1
    - starting at char position: 2619
  - found: 'pain'
    - on page 1
    - starting at char position: 2920
  - found: 'pain'
    - on page 1
    - starting at char position: 3164
  - found: 'pain'
    - on page 1
    - starting at char position: 3348
Found 15 results for "pain"
```

[Text Search and Replace]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html
[Finding the First Matching Phrase]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter09.html#finding-the-first-matching-phrase
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
