
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

Run from current example folder.

```shell
python start.py
```


## Output

```text
Loading Office...
No. of dictionaries: 3
  standard.dic (1); (active); ''; positive
  technical.dic (355); (active); ''; positive
  List of Ignored Words (0); (active); ''; positive

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
  IsSpellAuto: False
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
   Version: 1.2.8
   Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu59147xjqms4.tmp_/apso-v2.oxt

2. ID: French.linguistic.resources.from.Dicollecte.by.OlivierR
   Version: 5.7
   Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu287421qavj.tmp_/lo-oo-ressources-linguistiques-fr-v5-7.oxt

3. ID: org.openoffice.languagetool.oxt
   Version: 5.8
   Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu14553844wbl51.tmp_/LanguageTool-stable.oxt

4. ID: mytools.mri
   Version: 1.3.3
   Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu1050215332vj9.tmp_/MRI-1.3.3.oxt

5. ID: spanish.es.dicts.from.rla-es
   Version: 2.6
   Loc: file:///home/user/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu305561ujwy.tmp_/es.oxt

6. ID: org.openoffice.legacy.nlpsolver
   Version: 
   Loc: file:///usr/lib/libreoffice/share/extensions/nlpsolver

7. ID: org.openoffice.legacy.wiki-publisher
   Version: 
   Loc: file:///usr/lib/libreoffice/share/extensions/wiki-publisher

Available Services:
SpellChecker (1):
  org.openoffice.lingu.MySpellSpellChecker
Thesaurus (1):
  org.openoffice.lingu.new.Thesaurus
Hyphenator (1):
  org.openoffice.lingu.LibHnjHyphenator
Proofreader (1):
  org.languagetool.openoffice.Main
Configured Services:
SpellChecker (1):
  org.openoffice.lingu.MySpellSpellChecker
Thesaurus (1):
  org.openoffice.lingu.new.Thesaurus
Hyphenator (1):
  org.openoffice.lingu.LibHnjHyphenator
Proofreader (1):
  org.languagetool.openoffice.Main

Locales for SpellChecker (39)

    AR  BE  BF  BJ  BO  CA  CA  CH  CI
  CL  CO  CR  CU  DO  EC  ES  FR  GQ  GT
  HN  LU  MC  ML  MX  NE  NI  PA  PE  PH
  PR  PY  SN  SV  TG  US  US  UY  VE

Locales for Thesaurus (38)

    AR  BE  BF  BJ  BO  CA  CH  CI  CL
  CO  CR  CU  DO  EC  ES  FR  GQ  GT  HN
  LU  MC  ML  MX  NE  NI  PA  PE  PH  PR
  PY  SN  SV  TG  US  US  UY  VE

Locales for Hyphenator (40)

      AR  BE  BF  BJ  BO  CA  CA  CH
  CI  CL  CO  CR  CU  DO  EC  ES  FR  GQ
  GT  HN  LU  MC  ML  MX  NE  NI  PA  PE
  PH  PR  PY  SN  SV  TG  US  US  UY  VE

Locales for Proofreader (106)

            AE  AF  AO  AR  AT
  AU  BE  BE  BE  BH  BO  BR  BY  CA  CA
  CD  CH  CH  CH  CI  CL  CM  CN  CR  CU
  CV  DE  DE  DK  DO  DZ  EC  EG  ES  ES
  ES  ES  ES  FI  FR  FR  GB  GR  GT  GW
  HN  HT  IE  IN  IQ  IR  IT  JO  JP  KH
  KW  LB  LI  LU  LU  LY  MA  MA  MC  ML
  MO  MX  MZ  NI  NL  NZ  OM  PA  PE  PH
  PL  PR  PT  PY  QA  RE  RO  RU  SA  SD
  SE  SI  SK  SN  ST  SV  SY  TL  TN  UA
  US  US  UY  VE  YE  ZA


* 'ceurse' is unknown. Try:
No. of names: 3
  'ceruse'  'course'  'curse'


* 'magisian' is unknown. Try:
No. of names: 2
  'magician'  'siamang'


* 'ellucidate' is unknown. Try:
No. of names: 3
  'elucidate'  'elucidation'  'elucidator'


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
