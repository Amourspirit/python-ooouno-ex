
<p align="center">
<img src="https://user-images.githubusercontent.com/4193389/185815096-19db1273-709e-4f1c-b4ac-e4090183ac9b.png" alt="logo"/>
</p>

# Lingo

Examples using the Linguistics API (the `com.sun.star.linguistic2` module).
See Dev Guide, ch.6, p.706, "Linguistics".
<https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Linguistics>

Use the spell checker, thesaurus, proof reader (grammar checker), and
language guesser. The `hypernator` isn't used.

The proof reader uses `LanguageTool` (<https://www.languagetool.org/>)
not the default LightProof on one of my test machines.
(<http://extensions.libreoffice.org/extension-center/lightproof-editor>;
<http://libreoffice.hu/2011/12/08/grammar-checking-in-libreoffice/>)

## See

See Also:

- [The Linguistics API]
- [OOO Development Tools]

See [source code](./start.py)

## Automate

### Cross Platform

From this folder.

```sh
python start.py
```

### Linux/Mac

From project root folder

```sh
python ./ex/auto/writer/odev_lingo/start.py
```

### Windows

From project root folder

```ps
python .\ex\auto\writer\odev_lingo\start.py
```

## Live LibreOffice Python

Instructions to run this example in [Live-LibreOffice-Python](https://github.com/Amourspirit/live-libreoffice-python).

Start Live-LibreOffice-Python in a Codespace or in a Dev Container.

In the terminal run:

```bash
cd examples
gitget 'https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_lingo'
```

This will copy the `odev_lingo` example to the examples folder.

In the terminal run:

```bash
cd odev_lingo
python -m start
```

## Output

```text
Loading Office...
No. of dictionaries: 5
  standard.dic (1); (active); ''; positive
  en-GB.dic (42); (active); 'GB'; positive
  en-US.dic (42); (active); 'US'; positive
  technical.dic (355); (active); ''; positive
  List of Ignored Words (2); (active); ''; positive

No. of conversion dictionaries: 0

Linguistic Manager Properties
  DefaultLanguage: 0
  DefaultLocale: (com.sun.star.lang.Locale){ Language = (string)"", Country = (string)"", Variant = (string)"" }
  DefaultLocale_CJK: (com.sun.star.lang.Locale){ Language = (string)"", Country = (string)"", Variant = (string)"" }
  DefaultLocale_CTL: (com.sun.star.lang.Locale){ Language = (string)"", Country = (string)"", Variant = (string)"" }
  HyphMinLeading: 2
  HyphMinTrailing: 2
  HyphMinWordLength: 5
  IsGermanPreReform: None
  IsHyphAuto: False
  IsHyphSpecial: True
  IsIgnoreControlCharacters: True
  IsSpellAuto: True
  IsSpellCapitalization: True
  IsSpellHide: None
  IsSpellInAllLanguages: None
  IsSpellSpecial: True
  IsSpellUpperCase: True
  IsSpellWithDigits: False
  IsUseDictionaryList: True
  IsWrapReverse: False


Extensions:
1. ID: apso.python.script.organizer
   Version: 1.3.0
   Loc: file:///C:/Users/bigby/AppData/Roaming/LibreOffice/4/user/uno_packages/cache/uno_packages/lu1271241oyk.tmp_/apso.oxt

2. ID: org.openoffice.en.hunspell.dictionaries
   Version: 2021.11.01
   Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/dict-en

3. ID: French.linguistic.resources.from.Dicollecte.by.OlivierR
   Version: 7.0
   Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/dict-fr

4. ID: org.openoffice.languagetool.oxt
   Version: 5.8
   Loc: file:///C:/Users/bigby/AppData/Roaming/LibreOffice/4/user/uno_packages/cache/uno_packages/lu107803j3h0.tmp_/LanguageTool-stable.oxt

5. ID: com.sun.star.comp.Calc.NLPSolver
   Version: 0.9
   Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/nlpsolver

6. ID: spanish.es.dicts.from.rla-es
   Version: __VERSION__
   Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/dict-es

7. ID: com.sun.wiki-publisher
   Version: 1.2.0
   Loc: file:///C:/Program%20Files/LibreOffice/program/../share/extensions/wiki-publisher

Available Services:
SpellChecker (1):
  org.openoffice.lingu.MySpellSpellChecker
Thesaurus (1):
  org.openoffice.lingu.new.Thesaurus
Hyphenator (1):
  org.openoffice.lingu.LibHnjHyphenator
Proofreader (2):
  org.languagetool.openoffice.Main
  org.libreoffice.comp.pyuno.Lightproof.en

Configured Services:
SpellChecker (1):
  org.openoffice.lingu.MySpellSpellChecker
Thesaurus (1):
  org.openoffice.lingu.new.Thesaurus
Hyphenator (1):
  org.openoffice.lingu.LibHnjHyphenator
Proofreader (1):
  org.libreoffice.comp.pyuno.Lightproof.en

Locales for SpellChecker (46)
  AR  AU  BE  BO  BS  BZ  CA  CA  CH  CL
  CO  CR  CU  DO  EC  ES  FR  GB  GH  GQ
  GT  HN  IE  IN  JM  LU  MC  MW  MX  NA
  NI  NZ  PA  PE  PH  PH  PR  PY  SV  TT
  US  US  UY  VE  ZA  ZW

Locales for Thesaurus (46)
  AR  AU  BE  BO  BS  BZ  CA  CA  CH  CL
  CO  CR  CU  DO  EC  ES  FR  GB  GH  GQ
  GT  HN  IE  IN  JM  LU  MC  MW  MX  NA
  NI  NZ  PA  PE  PH  PH  PR  PY  SV  TT
  US  US  UY  VE  ZA  ZW

Locales for Hyphenator (46)
  AR  AU  BE  BO  BS  BZ  CA  CA  CH  CL
  CO  CR  CU  DO  EC  ES  FR  GB  GH  GQ
  GT  HN  IE  IN  JM  LU  MC  MW  MX  NA
  NI  NZ  PA  PE  PH  PH  PR  PY  SV  TT
  US  US  UY  VE  ZA  ZW

Locales for Proofreader (111)
  AE  AF  AO  AR  AT  AU  BE  BE  BE  BH
  BO  BR  BS  BY  BZ  CA  CA  CD  CH  CH
  CH  CI  CL  CM  CN  CR  CU  CV  DE  DE
  DK  DO  DZ  EC  EG  ES  ES  ES  ES  ES
  FI  FR  FR  GB  GH  GR  GT  GW  HN  HT
  IE  IE  IN  IN  IQ  IR  IT  JM  JO  JP
  KH  KW  LB  LI  LU  LU  LY  MA  MA  MC
  ML  MO  MX  MZ  NA  NI  NL  NZ  OM  PA
  PE  PH  PH  PL  PR  PT  PY  QA  RE  RO
  RU  SA  SD  SE  SI  SK  SN  ST  SV  SY
  TL  TN  TT  UA  US  US  UY  VE  YE  ZA
  ZW


* 'ceurse' is unknown. Try:
No. of names: 2
  'course'  'curse'


* 'magisian' is unknown. Try:
No. of names: 2
  'magician'  'magnesia'


* 'ellucidate' is unknown. Try:
No. of names: 2
  'elucidate'  'elucidation'


'magician' found in thesaurus; number of meanings: 2
1. Meaning: (noun) prestidigitator
 No. of  synonyms: 6
    prestidigitator
    conjurer
    conjuror
    illusionist
    performer (generic term)
    performing artist (generic term)

2. Meaning: (noun) sorcerer
 No. of  synonyms: 6
    sorcerer
    wizard
    necromancer
    thaumaturge
    thaumaturgist
    occultist (generic term)

Proofing...
G* This sentence does not start with an uppercase letter. in: 'i'
  Suggested change: 'I'

G* Spelling mistake in: 'dont'
  Suggested change: 'don't'

G* Word repetition in: 'one one'
  Suggested change: 'one'

No. of proofing errors: 3

Locale lang: 'en'; country: ''; variant: ''
Guessed language: en
Guessed language: fr
Closing Office
Office terminated
Office bridge has gone!!
```

[The Linguistics API]: https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/part2/chapter10.html
[OOO Development Tools]: https://python-ooo-dev-tools.readthedocs.io/en/latest/
