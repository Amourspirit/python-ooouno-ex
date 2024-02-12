from __future__ import annotations
from enum import Enum

import uno

from ooo.dyn.awt.font_weight import FontWeight
from ooo.dyn.chart.time_increment import TimeIncrement
from ooo.dyn.chart.time_interval import TimeInterval
from ooo.dyn.chart.time_unit import TimeUnit
from ooo.dyn.chart2.axis_orientation import AxisOrientation
from ooo.dyn.chart2.axis_type import AxisType
from ooo.dyn.drawing.line_style import LineStyle
from ooo.dyn.chart2.legend_position import LegendPosition
from ooodev.format.inner.direct.chart2.title.alignment.direction import (
    DirectionModeKind,
)

from ooodev.dialog.msgbox import (
    MessageBoxType,
    MessageBoxButtonsEnum,
    MessageBoxResultsEnum,
)
from ooodev.calc import CalcDoc, CalcSheet
from ooodev.office.chart2 import Chart2, Angle
from ooodev.office.chart2 import DataPointLabelTypeKind
from ooodev.office.chart2 import CurveKind
from ooodev.utils.color import CommonColor
from ooodev.utils.file_io import FileIO
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.kind.data_point_label_placement_kind import (
    DataPointLabelPlacementKind,
)
from ooodev.loader import Lo
from ooodev.utils.type_var import PathOrStr

from ooodev.calc.chart2.chart_doc import ChartDoc
from ooodev.format.inner.preset.preset_gradient import PresetGradientKind

# from ooodev.format.inner.direct.chart2.title.alignment.direction import DirectionModeKind


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

            chart_doc: None | ChartDoc = None
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
                Chart2.print_chart_types(chart_doc.component)

                template_names = Chart2.get_chart_templates(chart_doc.component)
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

    def _col_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draw a column chart;
        # uses "Sneakers Sold this Month" table
        range_addr = sheet.rng("A2:B8")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="C3",
            width=15,
            height=11,
            diagram_name=ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
        )
        sheet["A1"].goto()

        chart_doc = tbl_chart.chart_doc
        _ = chart_doc.set_title(sheet["A1"].value)
        _ = chart_doc.axis_x.set_title(sheet["A2"].value)
        y_axis_title = chart_doc.axis_y.set_title(sheet["B2"].value)
        y_axis_title.style_orientation(angle=90)
        chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _multi_col_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a multiple column chart: 2D and 3D
        # uses "States with the Most Colleges and Universities by Type"
        range_addr = sheet.rng("E15:G21")
        d_name = ChartTypes.Column.TEMPLATE_STACKED.COLUMN
        # d_name = ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_DEEP_3D
        # d_name = ChartTypes.Column.TEMPLATE_PERCENT.COLUMN_FLAT_3D

        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="A22",
            width=20,
            height=11,
            diagram_name=d_name,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A13"].goto()

        _ = chart_doc.set_title(sheet["E13"].value)
        _ = chart_doc.axis_x.set_title(sheet["E15"].value)
        y_axis_title = chart_doc.axis_y.set_title(sheet["F14"].value)
        y_axis_title.style_orientation(angle=90)
        chart_doc.first_diagram.view_legend(True)
        chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)

        # for the 3D versions
        # hide labels
        # Chart2.show_axis_label(chart_doc=chart_doc.component, axis_val=AxisKind.Z, idx=0, is_visible=False)
        # Chart2.set_chart_shape_3d(chart_doc=chart_doc.component, shape=DataPointGeometry3DEnum.CYLINDER)
        return chart_doc

    def _col_line_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a column+line chart
        # uses "States with the Most Colleges and Universities by Type"
        range_addr = sheet.rng("E15:G21")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="B3",
            width=20,
            height=11,
            diagram_name=ChartTypes.ColumnAndLine.TEMPLATE_STACKED.COLUMN_WITH_LINE,
        )
        chart_doc = tbl_chart.chart_doc
        sheet["A13"].goto()

        _ = chart_doc.set_title(sheet["E13"].value)
        _ = chart_doc.axis_x.set_title(sheet["E15"].value)
        y_axis_title = chart_doc.axis_y.set_title(sheet["F14"].value)
        y_axis_title.style_orientation(angle=90)
        chart_doc.first_diagram.view_legend(True)
        chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _bar_chart(self, sheet: CalcSheet) -> ChartDoc:
        # uses "Sneakers Sold this Month" table
        range_addr = sheet.rng("A2:B8")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="B3",
            width=15,
            height=11,
            diagram_name=ChartTypes.Bar.TEMPLATE_STACKED.BAR,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A1"].goto()

        _ = chart_doc.set_title(sheet["A1"].value)
        _ = chart_doc.axis_x.set_title(sheet["A2"].value)
        y_axis_title = chart_doc.axis_y.set_title(sheet["B2"].value)
        y_axis_title.style_orientation(angle=90)
        chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _pie_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draw a pie chart, with legend and subtitle;
        # uses "Top 5 States with the Most Elementary and Secondary Schools"
        range_addr = sheet.rng("E2:F8")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="B10",
            width=12,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_DONUT.PIE,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A1"].goto()
        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            title = chart_doc.set_title(sheet["E1"].value)
            # set the title font to bold.
            title.style_font_general(b=True)
            _ = chart_doc.first_diagram.set_title(sheet["F2"].value)
            chart_doc.first_diagram.view_legend(True)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
            legend = chart_doc.first_diagram.get_legend()
            if legend is not None:
                # add formatting to Legend
                legend.style_border_line(color=CommonColor.DARK_MAGENTA, width=0.5)
                # legend.style_area_color
                # turn off transparency so we can set the background
                legend.style_area_transparency_transparency(0)
                # set a gradient background from a preset.
                legend.style_area_gradient_from_preset(
                    preset=PresetGradientKind.DEEP_OCEAN
                )
                # set the font color to white so it shows up on the dark background.
                legend.style_font_effect(color=CommonColor.WHITE)

        return chart_doc

    def _pie_3d_chart(self, sheet: CalcSheet) -> ChartDoc:
        from ooodev.format.chart2.direct.series.data_series.options import Orientation

        # draws a 3D pie chart with rotation, label change
        # uses "Top 5 States with the Most Elementary and Secondary Schools"
        range_addr = sheet.rng("E2:F8")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="B10",
            width=12,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_3D.PIE_3D,
        )
        chart_doc = tbl_chart.chart_doc
        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            title = chart_doc.set_title(sheet["E1"].value)
            # set the title font to bold.
            title.style_font_general(b=True)
            _ = chart_doc.first_diagram.set_title(sheet["F2"].value)
            chart_doc.first_diagram.view_legend(True)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
            legend = chart_doc.first_diagram.get_legend()
            if legend is not None:
                # add formatting to Legend
                legend.style_border_line(color=CommonColor.DARK_MAGENTA, width=0.5)
                # legend.style_area_color
                # turn off transparency so we can set the background
                legend.style_area_transparency_transparency(0)
                # set a gradient background from a preset.
                legend.style_area_gradient_from_preset(
                    preset=PresetGradientKind.DEEP_OCEAN
                )
                # set the font color to white so it shows up on the dark background.
                legend.style_font_effect(color=CommonColor.WHITE)

        sheet["A1"].goto()

        # data series options vary by chart type. For this reason options are not baked into data_series class object.
        orient = Orientation(
            chart_doc=chart_doc.component, clockwise=False, angle=Angle(45)
        )
        Chart2.style_data_series(chart_doc=chart_doc.component, styles=[orient])

        props_lst = Chart2.get_data_points_props(chart_doc=chart_doc.component, idx=0)
        Lo.print(f"Number of data props in first data series: {len(props_lst)}")

        # change only the last data point to be bold white 14pt
        data_series_tpl = chart_doc.get_data_series()
        data_series = data_series_tpl[0]
        dp = data_series[0]
        dp.style_font_general(weight=FontWeight.BOLD, size=14, color=CommonColor.WHITE)
        return chart_doc

    def _donut_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a 3D donut chart with 2 rings
        # uses the "Annual Expenditure on Institutions" table
        range_addr = sheet.rng("A44:C50")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="D43",
            width=15,
            height=11,
            diagram_name=ChartTypes.Pie.TEMPLATE_DONUT.DONUT,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A48"].goto()
        _ = chart_doc.set_title(sheet["A43"].value)
        chart_doc.first_diagram.view_legend(True)
        subtitle = f'Outer: {sheet["B44"].value}\nInner: {sheet["C44"].value}'
        _ = chart_doc.first_diagram.set_title(subtitle)
        chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)

        return chart_doc

        # Chart2.set_data_point_labels(chart_doc=chart_doc, label_type=DataPointLabelTypeKind.CATEGORY)

    def _area_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws an area (stacked) chart;
        # uses "Trends in Enrollment in Public Schools in the US" table
        range_addr = sheet.rng("E45:G50")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="A52",
            width=16,
            height=11,
            diagram_name=ChartTypes.Area.TEMPLATE_STACKED.AREA,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A43"].goto()

        _ = chart_doc.set_title(sheet["E43"].value)
        _ = chart_doc.axis_x.set_title(sheet["E45"].value)
        y_title = chart_doc.axis_y.set_title(sheet["F44"].value)
        y_title.style_orientation(angle=90)
        chart_doc.first_diagram.view_legend(True)
        chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _line_chart(self, sheet: CalcSheet) -> None:
        # draw a line chart with data points, no legend;
        # uses "Humidity Levels in NY" table

        range_addr = sheet.rng("A14:B21")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="D13",
            width=16,
            height=9,
            diagram_name=ChartTypes.Line.TEMPLATE_SYMBOL.LINE_SYMBOL,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A1"].goto()
        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            _ = chart_doc.set_title(sheet["A13"].value)
            _ = chart_doc.axis_x.set_title(sheet["A14"].value)
            _ = chart_doc.axis_y.set_title(sheet["B14"].value)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)

    def _lines_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draw a line chart with two lines;
        # uses "Trends in Expenditure Per Pupil"
        range_addr = sheet.rng("E27:G39")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="A40",
            width=22,
            height=11,
            diagram_name=ChartTypes.Line.TEMPLATE_SYMBOL.LINE_SYMBOL,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A39"].goto()

        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            _ = chart_doc.set_title(sheet["E26"].value)
            _ = chart_doc.axis_x.set_title(sheet["E27"].value)
            y_axis_title = chart_doc.axis_y.set_title("Expenditure per Pupil")
            y_axis_title.style_orientation(angle=90)

            # too crowded for data points
            chart_doc.set_data_point_labels(label_type=DataPointLabelTypeKind.NONE)

            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)

        return chart_doc

    def _scatter_chart(self, sheet: CalcSheet) -> ChartDoc:
        # data from http://www.mathsisfun.com/data/scatter-xy-plots.html;
        # uses the "Ice Cream Sales vs Temperature" table
        range_addr = sheet.rng("A110:B122")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="C109",
            width=16,
            height=11,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_SYMBOL,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A126"].goto()

        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            title = chart_doc.set_title(sheet["A109"].value)
            # set font to bold and blue
            title.style_font_general(b=True, color=CommonColor.DARK_BLUE)
            _ = chart_doc.axis_x.set_title(sheet["A110"].value)
            y_axis_title = chart_doc.axis_y.set_title(sheet["B110"].value)
            y_axis_title.style_orientation(angle=90)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)

            chart_doc.calc_regressions()
            _ = chart_doc.draw_regression_curve(curve_kind=CurveKind.LINEAR)
            return chart_doc

    def _scatter_line_log_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draw a x-y scatter chart using log scaling
        # uses the "Power Function Test" table
        range_addr = sheet.rng("E110:F120")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="A121",
            width=20,
            height=11,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_LINE_SYMBOL,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A121"].goto()
        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            _ = chart_doc.set_title(sheet["E109"].value)
            _ = chart_doc.axis_x.set_title(sheet["E110"].value)
            y_axis_title = chart_doc.axis_y.set_title(sheet["F110"].value)
            y_axis_title.style_orientation(angle=90)

            # change x- and y- axes to log scaling
            chart_doc.axis_x.scale(CurveKind.LOGARITHMIC)
            chart_doc.axis_y.scale(CurveKind.LOGARITHMIC)
            chart_doc.draw_regression_curve(curve_kind=CurveKind.POWER)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _scatter_line_error_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws an x-y scatter chart with lines and y-error bars
        # uses the smaller "Impact Data - 1018 Cold Rolled" table
        # the example comes from "Using Descriptive Statistics.pdf"

        range_addr = sheet.rng("A142:B146")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="F115",
            width=14,
            height=16,
            diagram_name=ChartTypes.XY.TEMPLATE_LINE.SCATTER_LINE_SYMBOL,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A123"].goto()
        with chart_doc:
            # Lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            _ = chart_doc.set_title(sheet["A141"].value)
            _ = chart_doc.axis_x.set_title(sheet["A142"].value)
            y_axis_title = chart_doc.axis_y.set_title("Impact Energy (Joules)")
            y_axis_title.style_orientation(angle=90)

            Lo.print("Adding y-axis error bars")
            error_label = f"{sheet.name}.C142"
            error_range = f"{sheet.name}.C143:C146"
            _ = chart_doc.set_y_error_bar(
                data_label=error_label, data_range=error_range
            )

            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _labeled_bubble_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a bubble chart with labels;
        # uses the "World data" table
        range_addr = sheet.rng("H63:J93")
        with sheet.calc_doc:
            # lock calc controllers while inserting chart
            tbl_chart = sheet.charts.insert_chart(
                rng_obj=range_addr,
                cell_name="A62",
                width=18,
                height=11,
                diagram_name=ChartTypes.Bubble.TEMPLATE_BUBBLE.BUBBLE,
            )

        sheet["A62"].goto()
        chart_doc = tbl_chart.chart_doc

        with chart_doc:
            # lock chart controls while setting chart properties.
            # This is a little faster and avoids flicker
            _ = chart_doc.set_title(sheet["H62"].value)
            _ = chart_doc.axis_x.set_title(sheet["H63"].value)
            y_axis_title = chart_doc.axis_y.set_title(sheet["I63"].value)
            y_axis_title.style_orientation(angle=90)
            chart_doc.first_diagram.view_legend(True)

            # change the data points
            ds_arr = chart_doc.get_data_series()
            ds = ds_arr[0]
            ds.transparency = 50
            ds.border_style = LineStyle.SOLID
            ds.border_color = CommonColor.RED
            ds.label_placement = DataPointLabelPlacementKind.CENTER.value

            # Chart2.set_data_point_labels(chart_doc=chart_doc, label_type=DataPointLabelTypeKind.NUMBER)

            label = f"{sheet.name}.K63"
            names = f"{sheet.name}.K64:K93"
            chart_doc.add_cat_labels(data_label=label, data_range=names)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _net_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a net chart;
        # uses the "No of Calls per Day" table
        range_addr = sheet.rng("A56:D63")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="E55",
            width=16,
            height=11,
            diagram_name=ChartTypes.Net.TEMPLATE_LINE.NET_LINE,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["E55"].goto()

        with chart_doc:
            # Lock chart controllers while setting chart properties.
            # This is a little faster and avoids flicker.
            _ = chart_doc.set_title(sheet["A55"].value)
            chart_doc.first_diagram.view_legend(True)
            chart_doc.set_data_point_labels(label_type=DataPointLabelTypeKind.NONE)

            # reverse x-axis so days increase clockwise around net
            sd = chart_doc.axis_x.get_scale_data()
            sd.Orientation = AxisOrientation.REVERSE
            chart_doc.axis_x.set_scale_data(sd)
            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _happy_stock_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a fancy stock chart
        # uses the "Happy Systems (HASY)" table

        range_addr = sheet.rng("A86:F104")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="A105",
            width=25,
            height=14,
            diagram_name=ChartTypes.Stock.TEMPLATE_VOLUME.STOCK_VOLUME_OPEN_LOW_HIGH_CLOSE,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A105"].goto()

        with chart_doc:
            # Lock chart controllers while setting chart properties.
            # This is a little faster and avoids flicker
            title = chart_doc.set_title(sheet["A85"].value)
            # set the title font to bold.
            title.style_font_general(b=True)
            _ = chart_doc.axis_x.set_title(sheet["A86"].value)
            y_axis_title = chart_doc.axis_y.set_title(sheet["B86"].value)
            y_axis_title.style_orientation(angle=90)
            y_axis2 = chart_doc.axis2_y
            if y_axis2 is None:
                raise ValueError("No secondary y-axis found")
            y2_title = y_axis2.set_title("Stock Value")
            y2_title.style_orientation(angle=90)

            chart_doc.set_data_point_labels(label_type=DataPointLabelTypeKind.NONE)
            # Chart2.view_legend(chart_doc=chart_doc, is_visible=True)

            # change 2nd y-axis min and max; default is poor ($0 - $20)

            sd = y_axis2.get_scale_data()
            # Chart2.print_scale_data("Secondary Y-Axis", sd)
            sd.Minimum = 83
            sd.Maximum = 103
            y_axis2.set_scale_data(sd)

            # change x-axis type from number to date
            sd = chart_doc.axis_x.get_scale_data()
            sd.AxisType = AxisType.DATE

            # set major increment to 3 days
            ti = TimeInterval(Number=3, TimeUnit=TimeUnit.DAY)
            tc = TimeIncrement()
            tc.MajorTimeInterval = ti
            sd.TimeIncrement = tc
            chart_doc.axis_x.set_scale_data(sd)

            # change color of "WhiteDay" and "BlackDay" block colors
            ct = ChartTypes.Stock.NAMED.CANDLE_STICK_CHART
            candle_ct = chart_doc.find_chart_type(chart_type=ct)
            # Props.show_obj_props("Stock chart", candle_ct)
            candle_ct.color_stock_bars(
                white_day_color=CommonColor.GREEN, black_day_color=CommonColor.RED
            )

            # thicken the high-low line; make it yellow
            ds = chart_doc.get_data_series(chart_type=ct)
            Lo.print(f"No. of data series in candle stick chart: {len(ds)}")
            ds[0].line_width = 120  # LineWidth in 1/100 mm
            ds[0].color = CommonColor.YELLOW

            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _stock_prices_chart(self, sheet: CalcSheet) -> ChartDoc:
        # draws a stock chart, with an extra pork bellies line
        range_addr = sheet.rng("E141:I146")
        tbl_chart = sheet.charts.insert_chart(
            rng_obj=range_addr,
            cell_name="E148",
            width=12,
            height=11,
            diagram_name=ChartTypes.Stock.TEMPLATE_VOLUME.STOCK_OPEN_LOW_HIGH_CLOSE,
        )
        chart_doc = tbl_chart.chart_doc

        sheet["A150"].goto()

        with chart_doc:
            # Lock chart controllers while setting chart properties.
            # This is a little faster and avoids flicker
            title = chart_doc.set_title(sheet["E140"].value)
            # set the title font to bold.
            title.style_font_general(b=True)
            chart_doc.set_data_point_labels(label_type=DataPointLabelTypeKind.NONE)
            _ = chart_doc.axis_x.set_title(sheet["E141"].value)
            y_axis_title = chart_doc.axis_y.set_title("Dollars")
            y_axis_title.style_orientation(angle=90)

            Lo.print("Adding Pork Bellies line")
            pork_label = f"{sheet.name}.J141"
            pork_points = f"{sheet.name}.J142:J146"
            chart_doc.add_stock_line(data_label=pork_label, data_range=pork_points)
            chart_doc.first_diagram.view_legend(True)
            legend = chart_doc.first_diagram.get_legend()
            # move the legend to the bottom of the chart
            if legend is not None:
                legend.style_position(pos=LegendPosition.PAGE_END, no_overlap=True)

            chart_doc.style_border_line(color=CommonColor.DARK_BLUE, width=0.8)
        return chart_doc

    def _default_chart(self, sheet: CalcSheet) -> ChartDoc:
        # create a chart by using Chart2 defaults.
        # uses "Sneakers Sold this Month" table
        _ = sheet.select_cells_addr("B3:B7")
        Chart2.insert_chart()
        tbl_chart = sheet.charts[0]
        chart_doc = tbl_chart.chart_doc

        sheet.deselect_cells()
        sheet["A3"].goto()
        chart_doc.first_diagram.view_legend(True)
        return chart_doc
