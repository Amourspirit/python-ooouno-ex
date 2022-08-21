#!/usr/bin/env python
# coding: utf-8
#
# Examples using the Linguistics API (the com.sun.star.linguistic2 module).
# See Dev Guide, ch.6, p.706, "Linguistics".
# <https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Linguistics>

# Use the spell checker, thesaurus, proof reader (grammar checker), and
# language guesser. The hypernator isn't used.

# The proof reader uses LanguageTool (<https://www.languagetool.org/>)
# (<http://extensions.libreoffice.org/extension-center/lightproof-editor>;
# <http://libreoffice.hu/2011/12/08/grammar-checking-in-libreoffice/>)

from ooodev.office.write import Write
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props

from com.sun.star.linguistic2 import XLinguServiceManager2
from com.sun.star.linguistic2 import XSpellChecker


def main() -> int:

    with Lo.Loader(Lo.ConnectSocket(headless=True)) as loader:

        # print linguistics info
        Write.dicts_info()

        lingu_props = Write.get_lingu_properties()
        Props.show_props("Linguistic Manager", lingu_props)

        Info.list_extensions()  # these include linguistic extensions

        lingo_mgr = Lo.create_instance_mcf(XLinguServiceManager2, "com.sun.star.linguistic2.LinguServiceManager")
        if lingo_mgr is None:
            print("No linguistics manager found")
            return 0

        Write.print_services_info(lingo_mgr)

        # load spell checker
        speller = Lo.create_instance_mcf(XSpellChecker, "com.sun.star.linguistic2.SpellChecker")
        # it is possible to age a spell checker from lingo_mgr;
        # however, it results in a error when passed to Write.spell_word for unknow reason.
        # For this reason we go wiht the speller above and not the next line.
        # speller = lingo_mgr.getSpellChecker()

        # load thesaurus
        thesaurus = lingo_mgr.getThesaurus()

        # use spell checker
        Write.spell_word("horseback", speller)
        Write.spell_word("ceurse", speller)
        #  Write.spellWord("CEURSE", speller)
        Write.spell_word("magisian", speller)
        Write.spell_word("ellucidate", speller)

        Write.print_meaning("magician", thesaurus)

        # load & use proof reader (Lightproof or LanguageTool)
        proofreader = Write.load_proofreader()
        print("Proofing...")
        num_errs = Write.proof_sentence("i dont have one one dogs.", proofreader)
        print(f"No. of proofing errors: {num_errs}")
        print()

        # guess the language
        loc = Write.guess_locale("The rain in Spain stays mainly on the plain.")
        Write.print_locale(loc)

        if loc is not None:
            print("Guessed language: " + loc.Language)

        loc = Write.guess_locale("A vaincre sans p�ril, on triomphe sans gloire.")

        if loc is not None:
            print("Guessed language: " + loc.Language)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
