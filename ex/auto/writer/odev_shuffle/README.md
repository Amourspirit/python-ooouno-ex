# Shuffle Words

Each word in the input file is mid-shuffled.
This causes the middle letters of the word to be rearranged, but not the first
and last letters. Words of <= 3 characters are unaffected.
The words are highlighted as they are shuffled.

In this example the file "shuffled.odt" is saved in the project root folder.

## See

See Also: [Text API Overview], [Inserting/Changing Text in a Document]

See [source code](./start.py)

## Automate

Run from current example folder.

```shell
python start.py --file "../../../../resources/odt/cicero_dummy.odt"
```


![shuffle text](https://user-images.githubusercontent.com/4193389/184251513-a8c96a5d-85b0-42ff-a891-ee5762e46a24.gif)

[Text API Overview]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html

[Inserting/Changing Text in a Document]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter05.html#inserting-changing-text-in-a-document
