from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_TICK_MARK, XL_TICK_LABEL_POSITION

from utilities.style_variables import get_brand_color, get_chart_style


def get_chart_type(chart_type, preferred_color='*BLUE'):
    chart_types = {
        'COLUMN': ColumnChartStyle,
        'STACKED COLUMN': StackedColumnChartStyle,
        '100% STACKED COLUMN': HundredPercentStackedColumnChartStyle,
        'PIE': PieChartStyle,
        # 'DONUT': DonutChartStyle,  Causes errors. Donut is too common a word in foodservice
        'BAR': BarChartStyle,
        'STACKED BAR': StackedBarChartStyle,
        '100% STACKED BAR': HundredPercentStackedBarChartStyle,
        'LINE': LineChartStyle,
        'TABLE': TableStyle,
        'STAT': None,
        'PICTURE': None
    }
    try:
        return chart_types.get(chart_type)(preferred_color)
    except TypeError:
        return None


class DefaultChartStyle:
    emphasis = ['OVERALL', 'SUM', 'TOTAL'],
    de_emphasis = ['OTHER', 'SAME', 'PREFER NOT TO SAY', 'NEVER', 'OTHER: PLEASE SPECIFY']
    legend_locations = {'bottom': XL_LEGEND_POSITION.BOTTOM, 'right': XL_LEGEND_POSITION.RIGHT}

    def __init__(self, preferred_color='*BLUE'):
        self.chart_type_code = None

        # Overall Styles
        self.font = 'Arial'
        self.font_color = 'BLACK'
        self.font_size = 11
        self.has_labels = True

        # Color
        self.preferred_color = preferred_color
        self.vary_category_color = False
        self.differing_data_labels = False

        # Legend
        self.legend_location = self.legend_locations.get('right')
        self.value_table_legend = False

        # Plot Overlap
        self.gap = 70
        self.overlap = -50

        # Category Axis
        self.c_axis_visible = True
        self.c_axis_reverse = False
        self.c_axis_major_gridlines = False
        self.c_axis_minor_gridlines = False
        self.c_axis_maximum = None
        self.c_axis_minimum = None
        self.c_axis_line_color = get_brand_color('GRAY')
        self.c_axis_label_position = XL_TICK_LABEL_POSITION.NEXT_TO_AXIS
        self.c_axis_major_tick_mark = XL_TICK_MARK.NONE

        # Value Axis
        self.v_axis_visible = False
        self.v_axis_major_gridlines = False
        self.v_axis_minor_gridlines = False
        self.v_axis_maximum = None
        self.v_axis_minimum = None
        self.v_axis_line_color = get_brand_color('GRAY')

    @property
    def chart_color(self):
        return get_chart_style(self.preferred_color)


class PieChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.PIE
        self.legend_location = self.legend_locations.get('bottom')
        self.differing_data_labels = True
        self.vary_category_color = True


class DonutChartStyle(PieChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.DOUGHNUT
        self.differing_data_labels = True
        self.vary_category_color = True


class BarChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.BAR_CLUSTERED
        self.c_axis_reverse = True


class StackedBarChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.BAR_STACKED
        self.legend_location = self.legend_locations.get('bottom')
        self.overlap = 100
        self.c_axis_reverse = True
        self.differing_data_labels = True


class HundredPercentStackedBarChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.BAR_STACKED
        self.legend_location = self.legend_locations.get('bottom')
        self.overlap = 100
        self.v_axis_maximum = 1
        self.c_axis_reverse = True
        self.differing_data_labels = True


class ColumnChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.COLUMN_CLUSTERED
        self.legend_location = self.legend_locations.get('bottom')


class StackedColumnChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.COLUMN_STACKED
        self.overlap = 100
        self.differing_data_labels = True


class HundredPercentStackedColumnChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.COLUMN_STACKED
        self.v_axis_maximum = 1
        self.overlap = 100
        self.differing_data_labels = True


class LineChartStyle(DefaultChartStyle):
    def __init__(self, preferred_color):
        super().__init__(preferred_color)
        self.chart_type_code = XL_CHART_TYPE.LINE
        self.v_axis_visible = True
        self.c_axis_label_position = XL_TICK_LABEL_POSITION.LOW
        self.c_axis_line_color = get_brand_color('BLACK')
        self.value_table_legend = True


class TableStyle:
    def __init__(self, preferred_color):
        self.chart_type_code = None
        self.h_banding = True
        self.v_banding = False

