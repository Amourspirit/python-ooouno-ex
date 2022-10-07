<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/186020677-bb548a86-3bf7-4b04-b0f9-8f6a32428e26.jpg" alt="logo"/>
</p>


# Lingo File

Apply the spell checker and proof reader (grammar checker) to the supplied file.

## See

See Also:

- [The Linguistics API]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From this folder.

```sh
python -m start --file "../../../../resources/odt/badGrammar.odt"
```

### Linux/Mac

From project root folder.

```sh
python ./ex/auto/writer/odev_lingo_file/start.py --file "resources/odt/badGrammar.odt"
```

### Windows

From project root folder.

```ps
python .\ex\auto\writer\odev_lingo_file\start.py --file "resources/odt/badGrammar.odt"
```

## Output

```text
>> I have a dogs. I have one dogs.
G* The plural noun “dogs” cannot be used with the article “a”. Did you mean “a dog” or “dogs”? in: 'a dogs'
  Suggested change: 'a dog'


>> I allow of of go home.  i dogs. John don’t like dogs. So recieve no cats also.
G* Word repetition in: 'of of'
  Suggested change: 'of'

G* The personal pronoun “I” should be uppercase. in: 'i'
  Suggested change: 'I'

* 'recieve' is unknown. Try:
No. of names: 3
  'receive'  'relieve'  'reverie'
```

[The Linguistics API]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter10.html
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/

