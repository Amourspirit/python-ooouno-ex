from __future__ import annotations
from enum import Enum

import uno
from com.sun.star.chart2 import XChartDocument
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import Calc, CalcDoc, CalcSheet
from ooodev.office.chart2 import Chart2, Angle, DataPointLabelTypeKind, CurveKind, mEx
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.gui import GUI
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.kind.data_point_lable_placement_kind import (
    DataPointLabelPlacementKind,
)
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr

from ooo.dyn.awt.font_weight import FontWeight
from ooo.dyn.chart.time_increment import TimeIncrement
from ooo.dyn.chart.time_interval import TimeInterval
from ooo.dyn.chart.time_unit import TimeUnit
from ooo.dyn.chart2.axis_orientation import AxisOrientation
from ooo.dyn.chart2.axis_type import AxisType
from ooo.dyn.drawing.line_style import LineStyle


class ChartKind(str, Enum):
    AREA = "area"
    BAR = "bar"
    BUBBLE_LABELED = "bubble_labeled"
    COLUMN = "col"
    COLUMN_LINE = "col_line"
    COLUMN_MULTI = "col_multi"
    DONUT = "donut"
    HAPPY_STOCK = "happy_stock"
    LINE = "line"
    LINES = "lines"
    NET = "net"
    PIE = "pie"
    PIE_3D = "pie_3d"
    SCATTER = "scatter"
    SCATTER_LINE_ERROR = "scatter_line_error"
    SCATTER_LINE_LOG = "scatter_line_log"
    STOCK_PRICES = "stock_prices"
    DEFAULT = "default"


class Chart2View:
    def __init__(self, data_fnm: PathOrStr, chart_kind: ChartKind) -> None:
        _ = FileIO.is_exist_file(data_fnm, True)
        self._data_fnm = FileIO.get_absolute_path(data_fnm)
        self._chart_kind = chart_kind

    def main(self) -> None:
        _ = Lo.load_office(connector=Lo.ConnectPipe(), opt=Lo.Options(verbose=True))

        try:
            doc = CalcDoc.open_doc(fnm=self._data_fnm, visible=True)
            sheet = doc.sheets[0]

            chart_doc = None
            if self._chart_kind == ChartKind.AREA:
                chart_doc = self._area_chart(sheet)
            elif self._chart_kind == ChartKind.BAR:
                chart_doc = self._bar_chart(sheet)
            elif self._chart_kind == ChartKind.BUBBLE_LABELED:
                chart_doc = self._labeled_bubble_chart(sheet)
            elif self._chart_kind == ChartKind.COLUMN:
                chart_doc = self._col_chart(sheet)
            elif self._chart_kind == ChartKind.COLUMN_LINE:
                chart_doc = self._col_line_chart(sheet)
            elif self._chart_kind == ChartKind.COLUMN_MULTI:
                chart_doc = self._multi_col_chart(sheet)
            elif self._chart_kind == ChartKind.DONUT:
                chart_doc = self._donut_chart(sheet)
            elif self._chart_kind == ChartKind.HAPPY_STOCK:
                chart_doc = self._happy_stock_chart(sheet)
            elif self._chart_kind == ChartKind.LINE:
                chart_doc = self._line_chart(sheet)
            elif self._chart_kind == ChartKind.LINES:
                chart_doc = self._lines_chart(sheet)
            elif self._chart_kind == ChartKind.NET:
                chart_doc = self._net_chart(sheet)
            elif self._chart_kind == ChartKind.PIE:
                chart_doc = self._pie_chart(sheet)
            elif self._chart_kind == ChartKind.PIE_3D:
                chart_doc = self._pie_3d_chart(sheet)
            elif self._chart_kind == ChartKind.SCATTER:
                chart_doc = self._scatter_chart(sheet)
            elif self._chart_kind == ChartKind.SCATTER_LINE_ERROR:
                chart_doc = self._scatter_line_error_chart(sheet)
            elif self._chart_kind == ChartKind.SCATTER_LINE_LOG:
                chart_doc = self._scatter_line_log_chart(sheet)
            elif self._chart_kind == ChartKind.STOCK_PRICES:
                chart_doc = self._stock_prices_chart(sheet)
            elif self._chart_kind == ChartKind.DEFAULT:
                chart_doc = self._default_chart(sheet)

            if chart_doc:
                Chart2.print_chart_types(chart_doc)

                template_names = Chart2.get_chart_templates(chart_doc)
                Lo.print_names(template_names, 1)

            Lo.delay(2000)
            msg_result = doc.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close()
                Lo.close_office()
            else:
                print("Keeping document open")
        except Exception:
            Lo.close_office()
            raise

    def _col_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draw a column chart;
        # uses "Sneakers Sold this Month" table
        range_addr = sheet.get_address(range_name="A2:B8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="C3",
            width=15,
            height=11,
            diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
        )
        sheet["A1"].goto()

        _ = Chart2.set_title(
            chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A1")
        )
        _ = Chart2.set_x_axis_title(
            chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="A2")
        )
        _ = Chart2.set_y_axis_title(
            chart_doc=chart_doc, title=Calc.get_string(sheet=sheet, cell_name="B2")
        )
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        return chart_doc

    def _multi_col_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a multiple column chart: 2D and 3D
        # uses "States with the Most Colleges and Universities by Type"
        range_addr = sheet.get_address(range_name="E15:G21")
        d_name = ChartTypes.Column.TEMPLATE_STACKED.COLUMN
        # d_name = ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_DEEP_3D
        # d_name = ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_FLAT_3D
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="A22",
            width=20,
            height=11,
            diagram_name=d_name,
        )
        ChartTypes.Column.TEMPLATE_STACKED.COLUMN
        sheet["A13"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E13"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["E15"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["F14"].get_string())
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

        # for the 3D versions
        # hide labels
        # Chart2.show_axis_label(chart_doc=chart_doc, axis_val=AxisKind.Z, idx=0, is_visible=False)
        # Chart2.set_chart_shape_3d(chart_doc=chart_doc, shape=DataPointGeometry3DEnum.CYLINDER)
        return chart_doc

    def _col_line_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a column+line chart
        # uses "States with the Most Colleges and Universities by Type"
        range_addr = sheet.get_address(range_name="E15:G21")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="B3",
            width=20,
            height=11,
            diagram_name=ChartTypes.ColumnAndLine.TEMPLATE_STACKED.COLUMN_WITH_LINE,
        )
        sheet["A13"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E13"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["E15"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["F14"].get_string())
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        return chart_doc

    def _bar_chart(self, sheet: CalcSheet) -> XChartDocument:
        # uses "Sneakers Sold this Month" table
        range_addr = sheet.get_address(range_name="A2:B8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="B3",
            width=15,
            height=11,
            diagram_name=ChartTypes.Bar.TEMPLATE_STACKED.BAR,
        )
        sheet["A1"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A1"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["A2"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["B2"].get_string())
        # rotate x-axis which is now the vertical axis
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        return chart_doc

    def _pie_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draw a pie chart, with legend and subtitle;
        # uses "Top 5 States with the Most Elementary and Secondary Schools"
        range_addr = sheet.get_address(range_name="E2:F8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="B10",
            width=12,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_DONUT.PIE,
        )
        sheet["A1"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E1"].get_string())
        Chart2.set_subtitle(chart_doc=chart_doc, subtitle=sheet["F2"].get_string())
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        return chart_doc

    def _pie_3d_chart(self, sheet: CalcSheet) -> XChartDocument:
        from ooodev.format.chart2.direct.series.data_series.options import Orientation

        # draws a 3D pie chart with rotation, label change
        # uses "Top 5 States with the Most Elementary and Secondary Schools"
        range_addr = sheet.get_address(range_name="E2:F8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="B10",
            width=12,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_3D.PIE_3D,
        )
        sheet["A1"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E1"].get_string())
        Chart2.set_subtitle(chart_doc=chart_doc, subtitle=sheet["F2"].get_string())
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

        # rotate around horizontal (x-axis) and vertical (y-axis)

        # for unknown reason in LibreOffice, version 7.6 this stopped working. only tested on Linux
        # diagram = chart_doc.getFirstDiagram()
        # Props.set(
        #     diagram,
        #     RotationHorizontal=0,  # -ve rotates bottom edge out of page; default is -60
        #     RotationVertical=-45,  # -ve rotates left edge out of page; default is 0 (i.e. no rotation)
        # )

        # this is a replacement for the code above this is no longer working
        orient = Orientation(chart_doc=chart_doc, clockwise=False, angle=Angle(45))
        Chart2.style_data_series(chart_doc=chart_doc, styles=[orient])

        # Props.show_obj_props("Diagram", diagram)

        # change all the data points to be bold white 14pt
        # ds = Chart2.get_data_series(chart_doc)
        # Lo.print(f"No. of data series: {len(ds)}")
        # Props.show_obj_props("Data Series 0", ds[0])
        # Props.set(ds[0], CharHeight=14.0, CharColor=CommonColor.WHITE, CharWeight=FontWeight.BOLD)

        props_lst = Chart2.get_data_points_props(chart_doc=chart_doc, idx=0)
        Lo.print(f"Number of data props in first data series: {len(props_lst)}")

        # change only the last data point to be bold white 14pt
        try:
            props = Chart2.get_data_point_props(
                chart_doc=chart_doc, series_idx=0, idx=0
            )
            Props.set(
                props,
                CharHeight=14.0,
                CharColor=CommonColor.WHITE,
                CharWeight=FontWeight.BOLD,
            )
        except mEx.NotFoundError:
            Lo.print("No Properties found for chart.")
        return chart_doc

    def _donut_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a 3D donut chart with 2 rings
        # uses the "Annual Expenditure on Institutions" table
        range_addr = sheet.get_address(range_name="A44:C50")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="D43",
            width=15,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_DONUT.DONUT,
        )
        sheet["A48"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A43"].get_string())
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        subtitle = f'Outer: {sheet["B44"].value}\nInner: {sheet["C44"].value}'
        Chart2.set_subtitle(chart_doc=chart_doc, subtitle=subtitle)
        return chart_doc

        # Chart2.set_data_point_labels(chart_doc=chart_doc, label_type=DataPointLabelTypeKind.CATEGORY)

    def _area_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws an area (stacked) chart;
        # uses "Trends in Enrollment in Public Schools in the US" table
        range_addr = sheet.get_address(range_name="E45:G50")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="A52",
            width=16,
            height=11,
            diagram_name=ChartTypes.Area.TEMPLATE_STACKED.AREA,
        )
        sheet["A43"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E43"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["E45"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["F44"].get_string())
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        return chart_doc

    def _line_chart(self, sheet: CalcSheet) -> None:
        # draw a line chart with data points, no legend;
        # uses "Humidity Levels in NY" table
        range_addr = sheet.get_address(range_name="A14:B21")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="D13",
            width=16,
            height=9,
            diagram_name=ChartTypes.Line.TEMPLATE_SYMBOL.LINE_SYMBOL,
        )
        sheet["A1"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A13"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["A14"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["B14"].get_string())

    def _lines_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draw a line chart with two lines;
        # uses "Trends in Expenditure Per Pupil"
        range_addr = sheet.get_address(range_name="E27:G39")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="A40",
            width=22,
            height=11,
            diagram_name=ChartTypes.Line.TEMPLATE_SYMBOL.LINE_SYMBOL,
        )
        sheet["A26"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E26"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["E27"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title="Expenditure per Pupil")
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

        # too crowded for data points
        Chart2.set_data_point_labels(
            chart_doc=chart_doc, label_type=DataPointLabelTypeKind.NONE
        )
        return chart_doc

    def _scatter_chart(self, sheet: CalcSheet) -> XChartDocument:
        # data from http://www.mathsisfun.com/data/scatter-xy-plots.html;
        # uses the "Ice Cream Sales vs Temperature" table
        range_addr = sheet.get_address(range_name="A110:B122")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="C109",
            width=16,
            height=11,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_SYMBOL,
        )
        sheet["A104"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A109"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["A110"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["B110"].get_string())
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

        Chart2.calc_regressions(chart_doc)

        Chart2.draw_regression_curve(chart_doc=chart_doc, curve_kind=CurveKind.LINEAR)
        return chart_doc

    def _scatter_line_log_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draw a x-y scatter chart using log scaling
        # uses the "Power Function Test" table
        range_addr = sheet.get_address(range_name="E110:F120")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="A121",
            width=20,
            height=11,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_LINE_SYMBOL,
        )
        sheet["A121"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E109"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["E110"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["F110"].get_string())
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

        # change x- and y- axes to log scaling
        _ = Chart2.scale_x_axis(chart_doc=chart_doc, scale_type=CurveKind.LOGARITHMIC)
        # Chart2.print_scale_data("x-axis", x_axis)
        _ = Chart2.scale_y_axis(chart_doc=chart_doc, scale_type=CurveKind.LOGARITHMIC)
        Chart2.draw_regression_curve(chart_doc=chart_doc, curve_kind=CurveKind.POWER)
        return chart_doc

    def _scatter_line_error_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws an x-y scatter chart with lines and y-error bars
        # uses the smaller "Impact Data - 1018 Cold Rolled" table
        # the example comes from "Using Descriptive Statistics.pdf"

        range_addr = sheet.get_address(range_name="A142:B146")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="F115",
            width=14,
            height=16,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_LINE_SYMBOL,
        )
        sheet["A123"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A141"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["A142"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title="Impact Energy (Joules)")
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

        Lo.print("Adding y-axis error bars")
        error_label = f"{sheet.name}.C142"
        error_range = f"{sheet.name}.C143:C146"
        Chart2.set_y_error_bars(
            chart_doc=chart_doc, data_label=error_label, data_range=error_range
        )
        return chart_doc

    def _labeled_bubble_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a bubble chart with labels;
        # uses the "World data" table
        range_addr = sheet.get_address(range_name="H63:J93")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="A62",
            width=18,
            height=11,
            diagram_name=ChartTypes.Bubble.TEMPLATE_BUBBLE.BUBBLE,
        )
        sheet["A62"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["H62"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["H63"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["I63"].get_string())
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

        # change the data points
        ds = Chart2.get_data_series(chart_doc)
        Props.set(
            ds[0],
            Transparency=50,
            BorderStyle=LineStyle.SOLID,
            BorderColor=CommonColor.RED,
            LabelPlacement=DataPointLabelPlacementKind.CENTER.value,
        )

        # Chart2.set_data_point_labels(chart_doc=chart_doc, label_type=DataPointLabelTypeKind.NUMBER)

        label = f"{sheet.name}.K63"
        names = f"{sheet.name}.K64:K93"
        Chart2.add_cat_labels(chart_doc=chart_doc, data_label=label, data_range=names)
        return chart_doc

    def _net_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a net chart;
        # uses the "No of Calls per Day" table
        range_addr = sheet.get_address(range_name="A56:D63")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="E55",
            width=16,
            height=11,
            diagram_name=ChartTypes.Net.TEMPLATE_LINE.NET_LINE,
        )
        sheet["E55"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A55"].get_string())
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        Chart2.set_data_point_labels(
            chart_doc=chart_doc, label_type=DataPointLabelTypeKind.NONE
        )

        # reverse x-axis so days increase clockwise around net
        x_axis = Chart2.get_x_axis(chart_doc)
        sd = x_axis.getScaleData()
        sd.Orientation = AxisOrientation.REVERSE
        x_axis.setScaleData(sd)
        return chart_doc

    def _happy_stock_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a fancy stock chart
        # uses the "Happy Systems (HASY)" table

        range_addr = sheet.get_address(range_name="A86:F104")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="A105",
            width=25,
            height=14,
            diagram_name=ChartTypes.Stock.TEMPLATE_VOLUME.STOCK_VOLUME_OPEN_LOW_HIGH_CLOSE,
        )
        sheet["A105"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["A85"].get_string())
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["A86"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title=sheet["B86"].get_string())
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))
        Chart2.set_y_axis2_title(chart_doc=chart_doc, title="Stock Value")
        Chart2.rotate_y_axis2_title(chart_doc=chart_doc, angle=Angle(90))

        Chart2.set_data_point_labels(
            chart_doc=chart_doc, label_type=DataPointLabelTypeKind.NONE
        )
        # Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

        # change 2nd y-axis min and max; default is poor ($0 - $20)
        y_axis2 = Chart2.get_y_axis2(chart_doc)
        sd = y_axis2.getScaleData()
        # Chart2.print_scale_data("Secondary Y-Axis", sd)
        sd.Minimum = 83
        sd.Maximum = 103
        y_axis2.setScaleData(sd)

        # change x-axis type from number to date
        x_axis = Chart2.get_x_axis(chart_doc)
        sd = x_axis.getScaleData()
        sd.AxisType = AxisType.DATE

        # set major increment to 3 days
        ti = TimeInterval(Number=3, TimeUnit=TimeUnit.DAY)
        tc = TimeIncrement()
        tc.MajorTimeInterval = ti
        sd.TimeIncrement = tc
        x_axis.setScaleData(sd)

        # rotate the axis labels by 45 degrees
        # x_axis = Chart2.get_x_axis(chart_doc)
        # Props.set(x_axis, TextRotation=45)

        # Chart2.print_chart_types(chart_doc)

        # change color of "WhiteDay" and "BlackDay" block colors
        ct = ChartTypes.Stock.NAMED.CANDLE_STICK_CHART
        candle_ct = Chart2.find_chart_type(chart_doc=chart_doc, chart_type=ct)
        # Props.show_obj_props("Stock chart", candle_ct)
        Chart2.color_stock_bars(
            ct=candle_ct, w_day_color=CommonColor.GREEN, b_day_color=CommonColor.RED
        )

        # thicken the high-low line; make it yellow
        ds = Chart2.get_data_series(chart_doc=chart_doc, chart_type=ct)
        Lo.print(f"No. of data series in candle stick chart: {len(ds)}")
        # Props.show_obj_props("Candle Stick", ds[0])
        Props.set(
            ds[0], LineWidth=120, Color=CommonColor.YELLOW
        )  # LineWidth in 1/100 mm
        return chart_doc

    def _stock_prices_chart(self, sheet: CalcSheet) -> XChartDocument:
        # draws a stock chart, with an extra pork bellies line
        range_addr = sheet.get_address(range_name="E141:I146")
        chart_doc = Chart2.insert_chart(
            sheet=sheet.component,
            cells_range=range_addr,
            cell_name="E148",
            width=12,
            height=11,
            diagram_name=ChartTypes.Stock.TEMPLATE_VOLUME.STOCK_OPEN_LOW_HIGH_CLOSE,
        )
        sheet["A139"].goto()

        Chart2.set_title(chart_doc=chart_doc, title=sheet["E140"].get_string())
        Chart2.set_data_point_labels(
            chart_doc=chart_doc, label_type=DataPointLabelTypeKind.NONE
        )
        Chart2.set_x_axis_title(chart_doc=chart_doc, title=sheet["E141"].get_string())
        Chart2.set_y_axis_title(chart_doc=chart_doc, title="Dollars")
        Chart2.rotate_y_axis_title(chart_doc=chart_doc, angle=Angle(90))

        Lo.print("Adding Pork Bellies line")
        pork_label = f"{sheet.name}.J141"
        pork_points = f"{sheet.name}.J142:J146"
        Chart2.add_stock_line(
            chart_doc=chart_doc, data_label=pork_label, data_range=pork_points
        )

        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        return chart_doc

    def _default_chart(self, sheet: CalcSheet) -> XChartDocument:
        # create a chart by using Chart2 defaults.
        # uses "Sneakers Sold this Month" table
        _ = sheet.select_cells_addr("B3:B7")
        chart_doc = Chart2.insert_chart()
        # deselect cells.
        sheet.deselect_cells()
        sheet["A3"].goto()
        Chart2.view_legend(chart_doc=chart_doc, is_visible=True)
        return chart_doc
