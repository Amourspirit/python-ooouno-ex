
<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/184557139-c11d846b-d0c7-417e-ba86-9ed851552f7b.png" alt="big story screen shot"/>
</p>

# Story Creator

This example reads in a text file, apply a new paragraph style, header, page
numbers in footer, A4 page style, title, and subtitle, and saves as "bigStory.doc" (Word 97 format)
in the working folder.

if `--visible` is `True` then message boxes are displayed asking if you want to save and close document.

## See

See Also:

- [Text Styles]
- [Creating a new style]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From this folder.

```sh
python -m start --show --verbose  --file "../../../../resources/txt/scandal.txt"
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/writer/odev_story_creator/start.py --show --verbose --file "resources/txt/scandal.txt"
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_story_creator\start.py --show --verbose --file "resources/txt/scandal.txt"
```

## Output

```text
Loading Office...
Creating Office document swriter
Paragraph Styles
No. of names: 125
  ---------------------|----------------------|----------------------|----------------------
  Addressee            | Appendix             | Bibliography 1       | Bibliography Heading
  Caption              | Contents 1           | Contents 10          | Contents 2
  Contents 3           | Contents 4           | Contents 5           | Contents 6
  Contents 7           | Contents 8           | Contents 9           | Contents Heading
  Drawing              | Endnote              | Figure               | Figure Index 1
  Figure Index Heading | First line indent    | Footer               | Footer left
  Footer right         | Footnote             | Frame contents       | Hanging indent
  Header               | Header and Footer    | Header left          | Header right
  Heading              | Heading 1            | Heading 10           | Heading 2
  Heading 3            | Heading 4            | Heading 5            | Heading 6
  Heading 7            | Heading 8            | Heading 9            | Horizontal Line
  Illustration         | Index                | Index 1              | Index 2
  Index 3              | Index Heading        | Index Separator      | List
  List 1               | List 1 Cont.         | List 1 End           | List 1 Start
  List 2               | List 2 Cont.         | List 2 End           | List 2 Start
  List 3               | List 3 Cont.         | List 3 End           | List 3 Start
  List 4               | List 4 Cont.         | List 4 End           | List 4 Start
  List 5               | List 5 Cont.         | List 5 End           | List 5 Start
  List Contents        | List Heading         | List Indent          | Marginalia
  Numbering 1          | Numbering 1 Cont.    | Numbering 1 End      | Numbering 1 Start
  Numbering 2          | Numbering 2 Cont.    | Numbering 2 End      | Numbering 2 Start
  Numbering 3          | Numbering 3 Cont.    | Numbering 3 End      | Numbering 3 Start
  Numbering 4          | Numbering 4 Cont.    | Numbering 4 End      | Numbering 4 Start
  Numbering 5          | Numbering 5 Cont.    | Numbering 5 End      | Numbering 5 Start
  Object index 1       | Object index heading | Preformatted Text    | Quotations
  Salutation           | Sender               | Signature            | Standard
  Subtitle             | Table                | Table Contents       | Table Heading
  Table index 1        | Table index heading  | Text                 | Text body
  Text body indent     | Title                | User Index 1         | User Index 10
  User Index 2         | User Index 3         | User Index 4         | User Index 5
  User Index 6         | User Index 7         | User Index 8         | User Index 9
  User Index Heading   |                      |                      |


A Writer document
Saving the document in '/home/user/Python/ooouno_ex/ex/auto/writer/odev_story_creator/bigStory.doc'
Using format MS Word 97
Closing the document
Closing Office
Office terminated
Office bridge has gone!!
```

[Text Styles]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter06.html
[Creating a new style]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter06.html#creating-a-new-style
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
