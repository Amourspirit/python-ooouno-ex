from __future__ import annotations

import uno
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.calc import Calc, FormatterTable


class CalcFunctions:
    def main(self) -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:
            doc = Calc.create_doc(loader)
            sheet = Calc.get_sheet(doc=doc, index=0)
            # round
            print("ROUND result for 1.999 is: ", end="")
            print(Calc.call_fun("ROUND", 1.999))
            print()

            # sine of 30 degrees
            print("SIN result for 30 degrees is:", end="")
            print(f'{Calc.call_fun("SIN", Calc.call_fun("RADIANS", 30)):.3f}')

            # average function
            avg = float(Calc.call_fun("AVERAGE", 1, 2, 3, 4, 5))
            print(f"Average of the numbers is: {avg:.1f}")
            print()

            # slope function
            # https://tinyurl.com/2lf34yxq#slope
            # the slope function only seems to work if passed XCellRange
            arr = [[1.0, 2.0, 3.0], [3.0, 6.0, 9.0]]
            Calc.set_array(values=arr, sheet=sheet, name="A1")
            Lo.delay(500)
            xrng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
            yrng = Calc.get_cell_range(sheet=sheet, range_name="A2:C2")
            slope = float(Calc.call_fun("SLOPE", yrng, xrng))
            print(f"SLOPE of the line: {slope}")
            print()

            # zTest function
            arr = ((1.0, 2.0, 3.0),)
            Calc.set_array(values=arr, sheet=sheet, name="A1")
            Lo.delay(500)
            rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C1")
            res = float(Calc.call_fun("ZTEST", rng, 2.0))
            print(f"ZTEST result for data ((1,2,3),) and 2.0 is: {res}")

            # easter sunday function
            easter_sun = float(Calc.call_fun("EASTERSUNDAY", 2015))
            day = round(Calc.call_fun("DAY", easter_sun))
            month = round(Calc.call_fun("MONTH", easter_sun))
            year = round(Calc.call_fun("YEAR", easter_sun))
            print(f"Easter Sunday (d/m/y): {day}/{month}/{year}")
            print()

            # transpose a matrix
            arr = [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
            Calc.set_array(values=arr, sheet=sheet, name="A1")
            Lo.delay(500)
            rng = Calc.get_cell_range(sheet=sheet, range_name="A1:C3")
            trans_mat = Calc.call_fun("TRANSPOSE", rng)
            # add a little extra formatting
            fl = FormatterTable(format=(".1f", ">5"))
            Calc.print_array(trans_mat, fl)

            # sum two imaginary numbers: "13+4j" + "5+3j" returns 18+7j.
            sum = Calc.call_fun("IMSUM", "13+4j", "5+3j")
            print(f"13+4j + 5+3j: {sum}")
            print()

            # decimal to hex
            # DEC2HEX: DEC2HEX(100;4) returns "0064"
            hex4 = Calc.call_fun("DEC2HEX", 100, 4)
            print(f"100 to hex: {hex4}")
            print()

            # ROT13(Text)
            # ROT13: ROT13("hello") returns "uryyb"
            # uryyb
            arg = "hello"
            rot13 = Calc.call_fun("ROT13", arg)
            print(f'ROT13 of "{arg}": {rot13}')
            print()

            # Roman numbers
            # http://cs.stackexchange.com/questions/7777/is-the-language-of-roman-numerals-ambiguous
            roman = Calc.call_fun("ROMAN", 999)
            roman4 = Calc.call_fun("ROMAN", 999, 4)
            print(f"999 in Roman numerals: {roman} or {roman4}")
            print()

            # using ADDRESS
            cell_name = Calc.call_fun("ADDRESS", 2, 5, 4)  # row, column, abs
            print(f"Relative address for (5,2): {cell_name}")
            print()

            # prints over 500 names
            # print("Function Names")
            # Lo.print_names(Calc.get_function_names(), 6)

            Calc.print_function_info("EASTERSUNDAY")
            Calc.print_function_info("ROMAN")

            self._show_recent_functions()

            Lo.close(closeable=doc, deliver_ownership=False)

    def _show_recent_functions(self) -> None:
        recent_ids = Calc.get_recent_functions()
        if not recent_ids:
            return

        print(f"Recently used functions {len(recent_ids)}")
        for i in recent_ids:
            p = Calc.find_function(idx=i)
            print(f'  {Props.get_value(name="Name", props=p)}')

        print()
