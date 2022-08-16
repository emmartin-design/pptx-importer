from collections import deque
from calendar import month_abbr

from xlsx_handlers.excel_handler import get_workbook
from pptx_handlers.chart_definitions import get_chart_type
from pptx_handlers.text_handler import ParagraphInstance
from utilities.style_variables import get_chart_style
from utilities.preflight import get_error_dict
from utilities.utility_functions import (
    concat_dfs,
    combine_lists,
    country_list,
    format_as_float,
    format_date,
    get_column_names,
    get_current_time,
    get_df_from_dict,
    get_df_from_worksheet,
    get_key_with_matching_parameters,
    is_null,
    pivot_dataframe,
    replace_chars,
    t_round,
    trim_df
)


def get_outline(report_type, excel_file, has_page_tags=False, **kwargs):
    """
    Each report type must have an outline class for this to function
    Report type must match report_classes key exactly
    """
    report_classes = {
        'General Import': GeneralReportData,
        'Global Navigator Country Reports': GeneralReportData,
        'LTO Scorecard Report': LTOScorecardReportData,
        'DirecTV Scorecard': DirecTVReportData,
        'Quarterly Consumer KPIs': ConsumerKPIReportData,
        # 'Value Scorecard Report': None, -- Retired
        # 'C-Store Consumer KPIs': None,  -- Retired
        # 'Subway Scorecard': None -- Retired
    }
    return report_classes.get(report_type)(excel_file, has_page_tags=has_page_tags, **kwargs)


def is_all_strings(lst: list):
    lst = [isinstance(x, str) for x in lst]
    return all(lst)


def is_sequential(lst: list):
    """
    Checks if these are sequential/scale category titles
    """
    return all([x == y for x, y in zip(lst, range(1, 1000))])


def remove_parentheticals(value):
    new_value = value
    if any([x in new_value for x in ['(', ')']]):
        new_value = new_value.replace(value[value.find('('):(value.find(')') + 1)], '')
    new_value = replace_chars(new_value, (" :'", ""), (" '", ""), (' : ', ''), ('NumericQuestion', ''))
    new_value = new_value.strip()
    return new_value


def reformat_vertical_series(data: {}):
    """
    If multiple vertical-series were selected a key "categories_1" should exist
    This will multiply "categories_0" to match its length for future pivots
    """
    if data.get('categories_1') is not None:
        if len(data.keys()) > 3:
            data.pop('categories_1')
        else:
            pivot_len = len(set(data['categories_1']))
            data['categories_0'] = combine_lists([[x] * pivot_len for x in data['categories_0']])
    return data


class ShapeMeta:
    def __init__(
            self,
            height,
            width,
            top,
            left,
            text=None,
            fill_color=None,
            line=None,
            line_color=None
    ):
        self.height = height
        self.width = width
        self.top = top
        self.left = left
        self.text = [] if text is None else text
        self.fill_color = fill_color
        self.line = line
        self.line_color = line_color


class ReportData:
    """
    Basic parent class for opening the excel file
    """

    def __init__(self, excel_file, log_prefix=None, has_page_tags=False, **kwargs):
        self.excel_file = excel_file
        self.wb = get_workbook(excel_file)
        self.has_page_tags = has_page_tags
        self.pages = []

    def get_report_data(self):
        pass


class PPTXChartMeta:
    def __init__(self, df, intended_chart_type, preferred_color='*BLUE'):
        self.has_data = True
        self.chart_title = None
        self.chart_question = None
        self.worksheet_name = ', '.join(df.index.tolist())
        self.df = df

        # Chart data
        self.number_format = '0%'
        self.decimal_places = 0

        # Chart config
        self.preferred_color = preferred_color
        self.highlight = False
        self.heat_map = False
        self.growth = False
        self.top_box = False
        self.transpose = False
        self.intended_chart_type = intended_chart_type.upper()
        self.emphasis = ['OVERALL', 'SUM', 'TOTAL', 'GLOBAL AVERAGE']
        self.de_emphasis = ['OTHER', 'SAME', 'PREFER NOT TO SAY', 'NEVER', 'OTHER: PLEASE SPECIFY']

        self.label_text = {}
        self.bases = []
        self.base_col_idxs = []
        self.copy = []

        self.chart_style = get_chart_type(self.intended_chart_type, self.preferred_color)
        self.indexes = [[-99999, 99999] for _ in self.df.columns]
        self.change_from_previous = None
        self.notes = []

    @property
    def excel_number_format(self):
        parameters = {
            f"0{'.' if self.decimal_places > 0 else ''}{'0' * self.decimal_places}%": ['%' in self.number_format],
            '0.0': [self.number_format == '0.0'],
            '#,##0': [self.number_format == '0'],
            '$#,0.00': [self.number_format == '$0.00']
        }
        return get_key_with_matching_parameters(parameters)


class PPTXFlexibleChartMeta:
    """
    Houses and formats data and metadata from excel tabs
    Only one chart per tab is possible
    Only data points formatted as floats or percents will be picked up.
    """
    title_cell_color = [4, 'FF4472C4']  # Theme color index 4
    base_cell_color = [5, 'FFED7D31']  # Theme color index 4 or hex value
    data_cell_color = [7, 8, 9, 'FFFFC000']  # Mix of indices and hex

    def __init__(self, worksheet, worksheet_name, preferred_color='*BLUE', log_prefix=None, report_focus=None):
        print(f'Reading {worksheet_name} data')
        # Worksheet Metadata
        self.worksheet = worksheet
        self.worksheet_name = worksheet_name
        self.state = worksheet.sheet_state
        self.tab_color = worksheet.sheet_properties.tabColor
        self.log_prefix = log_prefix
        self.has_data = False
        self.first_value_col = 0

        # Chart data
        self.title_and_question = []
        self.report_focus = report_focus
        self.base_check = False
        self.percent_check = True
        self.number_format = '0%'
        self.decimal_places = 0
        self.intended_chart_type = 'BAR'
        self.preferred_color = preferred_color
        self.bases = []
        self.base_col_idxs = []
        self.copy = []
        self.label_text = {}
        self.notes = []
        self.cell_cache = None  # Holds the values of the previous cell while iterating
        self.cell_format_cache = None

        # Data config
        self.transpose = False
        self.sort = False
        self.force_percent = False
        self.mean = False
        self.median = False
        self.force_float = False
        self.force_currency = False
        self.force_int = False
        self.top_5 = False
        self.top_10 = False
        self.top_20 = False
        self.top_box = False
        self.heat_map = False
        self.highlight = False
        self.trim_axis = False
        self.growth = False
        self.emphasis = ['OVERALL', 'SUM', 'TOTAL', 'GLOBAL AVERAGE']
        self.de_emphasis = ['OTHER', 'SAME', 'PREFER NOT TO SAY', 'NEVER', 'OTHER: PLEASE SPECIFY']

        self.sort_by = None
        self.vertical_series = False

        self.raw_data = self.get_data_from_excel()
        self.clean_data = self.clean_up_data()
        try:
            self.df = get_df_from_dict(self.clean_data)
        except ValueError as e:
            for x, y in self.clean_data.items():
                print(f'{x}: LEN{len(y)} LST {y}', )
            raise ValueError(e)

        self.format_df()
        self.remove_command_conflicts()
        self.update_bases()
        self.indexes = [[-99999, 99999] for _ in self.df.columns]
        self.change_from_previous = None

        self.chart_style = get_chart_type(self.intended_chart_type, self.preferred_color)

    @staticmethod
    def command_check_styling(value):
        new_string = replace_chars(value, (' ', '_'), ('*', ''))
        return new_string.lower()

    @staticmethod
    def format_as_string(value):
        return str(value)

    @property
    def chart_title(self):
        return str(min(self.title_and_question, key=len)).strip()

    @property
    def chart_question(self):
        return str(max(self.title_and_question, key=len)).strip()

    @property
    def excel_number_format(self):
        parameters = {
            f"0{'.' if self.decimal_places > 0 else ''}{'0' * self.decimal_places}%": [self.percent_check],
            '0.0': [self.force_float, self.mean, self.median],
            '#,##0': [self.force_int],
            '$#,0.00': [self.force_currency]
        }
        return get_key_with_matching_parameters(parameters)

    def remove_command_conflicts(self):
        if any([self.mean, self.force_float, self.force_currency, self.force_int]):
            self.percent_check = False

    def update_bases(self):
        if self.report_focus is None:
            return
        idxs = [idx for idx, country in enumerate(country_list) if country not in ['Global average', self.report_focus]]
        self.bases = [base for idx, base in enumerate(self.bases) if idx not in idxs]
        self.base_col_idxs = [idx for idx, _ in enumerate(self.bases)]

    def drop_base(self):
        if 'Base' in self.df.index and all([x == 1.0 for x in self.df.loc['Base'].values.tolist()]):
            self.df = self.df.drop(index='Base')
        if 'Base' in self.df.columns and all([x == 1.0 for x in self.df['Base'].tolist()]):
            self.df = self.df.drop(columns='Base')
        self.df = self.df.rename(columns={'Base': 'Overall'})

    def format_indices_and_columns(self):
        self.df = self.df.set_index('categories_0')
        if self.report_focus is not None and self.report_focus in country_list:
            self.df = self.df[['Global average', self.report_focus]]
        try:
            self.df = pivot_dataframe(self.df, 'categories_1')
            self.df.index.name = None
            self.drop_base()
        except ValueError as e:
            index = self.df.index.tolist()
            duplicates = [x for x in index if index.count(x) == max([index.count(x) for x in index], key=index.count)]
            self.notes.append(f'Could not pivot vertical series: {e}:, {duplicates}')
            self.notes.append(self.df.to_string())
            self.df = get_df_from_dict(get_error_dict())

    def create_top_box(self):
        self.df[f"Top {len(self.df.columns)} Box"] = self.df.sum(axis=1)

        if self.intended_chart_type.upper() not in ['STACKED BAR', 'STACKED COLUMN']:
            self.intended_chart_type = 'STACKED BAR'

    def transpose_check(self):
        if '100%' in self.intended_chart_type and not self.transpose:
            rows_add_to_one = all([0.99 < sum(self.df.loc[x].values.tolist()) < 1.01 for x in self.df.index])
            columns_add_to_one = all([0.99 < sum(self.df[x].tolist()) < 1.01 for x in self.df.columns])
            self.transpose = columns_add_to_one and not rows_add_to_one

    def truncate_df(self):
        if any([self.top_5, self.top_10, self.top_20]):
            sort_by_column = self.df.columns[0] if self.sort_by is None else self.sort_by
            self.df = self.df.sort_values(by=sort_by_column, ascending=False) if self.sort else self.df
            truncate_index = 5 if self.top_5 else (10 if self.top_10 else 20)
            self.df = self.df.copy().iloc[0:truncate_index]

    def apply_commands(self):
        sort_by_column = self.df.columns[0] if self.sort_by is None else self.sort_by
        self.df = self.df.transpose() if self.transpose else self.df
        self.df = self.df.sort_values(by=sort_by_column, ascending=False) if self.sort else self.df
        self.df = self.df.apply(lambda row: (row / 100), axis=1) if self.force_percent else self.df
        self.df = self.df.apply(lambda row: (row * 100), axis=1) if self.force_int else self.df
        self.df = self.create_top_box() if self.top_box else self.df
        self.truncate_df()

    def format_df(self):
        self.format_indices_and_columns()
        self.transpose_check()
        self.apply_commands()

    @staticmethod
    def check_for_column_title(col_data, first_col_string_value):
        if len(col_data) > 0 and not isinstance(col_data[0], str):
            col_data = [first_col_string_value] + col_data
        return col_data

    @staticmethod
    def get_first_string_in_col(cell_value, cell_is_merged, first_string_value):
        try:
            """
            Catches numbers, usually bases, that aren't actually strings
            """
            int_tracker = first_string_value
            int(int_tracker)
            return first_string_value
        except (ValueError, TypeError):
            if all([
                first_string_value is None,
                isinstance(cell_value, str),
                not cell_is_merged
            ]):
                return cell_value
            return first_string_value

    def get_data_from_excel(self):
        """Iterates worksheet col by col, cell by cell, and checks color and values"""
        cell_data_by_col = []
        for col_idx, column in enumerate(self.worksheet.iter_cols()):
            col_data = []
            first_string_value = None  # Used to prevent column name selection omissions
            for cell_idx, cell in enumerate(column):
                cell_value = self.worksheet[cell.coordinate].value
                cell_is_merged = type(cell).__name__ == 'MergedCell'
                if cell_value is not None:
                    data_point = self.process_data_colors(cell, cell_value)
                    first_string_value = self.get_first_string_in_col(cell_value, cell_is_merged, first_string_value)
                    if data_point is not None:
                        col_data.append(data_point)
                self.cell_cache = cell_value
                self.cell_format_cache = self.worksheet[cell.coordinate].number_format
            col_data = self.check_for_column_title(col_data, first_string_value)
            cell_data_by_col.append(col_data)
        return cell_data_by_col

    def extract_bases(self, list_of_data, list_idx, vertical_series_modifier):
        try:
            base_value = [x for x in list_of_data if 'Base: ' in str(x)][0].replace('Base: ', '')
            self.bases.append(base_value)
            self.base_col_idxs.append((list_idx - vertical_series_modifier))
        except IndexError:
            pass

    def clean_up_data(self):
        """
        Receives a list of lists and converts it into a dictionary for later dataframe conversion
        """
        new_data = {}
        vertical_series_modifier = 1
        for list_idx, list_of_data in enumerate([x for x in self.raw_data if len(x) > 0]):
            if is_all_strings(list_of_data):
                list_of_data = [remove_parentheticals(x) for x in list_of_data]
                new_data[f'categories_{list_idx}'] = [x for x in list_of_data if 'BASE: ' not in x.upper()]
            else:
                first_string_instance = [x for x in list_of_data if isinstance(x, str) and 'Base: ' not in x][0]
                new_data[first_string_instance] = [x for x in list_of_data if isinstance(x, float)]
            self.extract_bases(list_of_data, list_idx, vertical_series_modifier)
            vertical_series_modifier = 2 if 'categories_1' in new_data.keys() else 1
        new_data = {x:y for x, y in new_data.items() if len(y) > 0}
        new_data = reformat_vertical_series(new_data)
        return new_data

    def set_command_to_true(self, cell_value):
        if cell_value.startswith('*'):
            setattr(self, self.command_check_styling(cell_value), True)

    def set_sort_by_column(self, cell_value):
        self.sort = True
        column_sort = replace_chars(cell_value, ('*SORT BY ', ''))
        self.sort_by = column_sort if 'COUNTRY' not in cell_value.upper() else self.log_prefix

    def set_preferred_color(self, cell_value):
        if cell_value in {'*BLUE', '*RED', '*GREEN', '*ORANGE', '*YELLOW', '*PURPLE', '*MULTI'}:
            self.preferred_color = get_chart_style(cell_value.upper())

    def set_chart_type(self, cell_value):
        self.intended_chart_type = cell_value.upper()

    def process_data_values(self, cell_value):
        cell_value = str(cell_value)
        parameters = {
            self.set_command_to_true: [hasattr(self, self.command_check_styling(cell_value))],
            self.set_sort_by_column: ['SORT BY' in cell_value.upper()],
            self.set_preferred_color: [get_chart_style(cell_value.upper()) is not None, 'MULTI' in cell_value.upper()],
            self.set_chart_type: [get_chart_type(cell_value.upper()) is not None],
        }
        try:
            get_key_with_matching_parameters(parameters)(cell_value)
        except TypeError:  # Cell doesn't contain important info
            pass

    def get_formatted_value(self, cell_value, cell_number_format):
        parameters = {
            format_as_float: [
                cell_number_format in ['0.0', '0.00', '0%', '0.0%'],
                all([x == '-' for x in [cell_value, self.cell_cache]]),
                cell_value == '-' and all([isinstance(self.cell_cache, int), '%' not in self.cell_format_cache]),
                cell_value == '*' and all([isinstance(self.cell_cache, int), '%' not in self.cell_format_cache])
            ],
            self.format_as_string: [cell_number_format in ['@', 'General']],
        }
        return get_key_with_matching_parameters(parameters)(cell_value)

    def process_data_colors(self, cell, cell_value):
        self.process_data_values(cell_value)
        cell_number_format = self.worksheet[cell.coordinate].number_format
        cell_color = cell.fill.start_color.index
        if cell_color in self.title_cell_color:
            self.title_and_question.append(cell_value)
        elif cell_color in self.base_cell_color:
            return f'Base: {cell_value}'
        elif cell_color in self.data_cell_color:
            self.has_data = True
            return self.get_formatted_value(cell_value, cell_number_format)


class PPTXPageMeta:
    def __init__(self, charts: list = None, function=None, title=None, footer: list = None):
        self.charts = [] if charts is None else charts
        self.function = function

        try:
            self.chart_count = self.get_chart_count()
            self.table_count = len([x for x in charts if 'TABLE' in x.intended_chart_type.upper()])
        except TypeError:
            self.chart_count = 0
            self.table_count = 0

        self.section_tag = None
        self.title = title
        self.copy = []
        self.shapes = []
        self.pictures = []
        self.footer = footer if footer is not None else self.get_footer()

    @property
    def notes(self):
        return {x.worksheet_name: x.notes for x in self.charts}

    def get_footer(self):
        footer = []
        for chart in self.charts:
            print(f'Processing {chart.worksheet_name}')
            if len(chart.bases) == 1:
                bases = [f'Base: {chart.bases[0]}']
            elif chart.transpose:
                index = chart.df.index.tolist()
                bases = [f'{index[idx]} base: {base}' for idx, base in zip(chart.base_col_idxs, chart.bases)]
            else:
                bases = [f'{chart.df.columns[idx]} base: {base}' for idx, base in zip(chart.base_col_idxs, chart.bases)]
            footer.append(', '.join(bases))
            if chart.chart_question is not None:
                footer.append(f'Q: {chart.chart_question}')
        return footer

    def get_chart_count(self):
        chart_count = 0
        for chart in self.charts:
            if not any([
                get_chart_type(chart.intended_chart_type.upper()) is None,
                chart.intended_chart_type.upper() == 'TABLE'
            ]):
                chart_count += 1
        return chart_count


class GeneralReportData(ReportData):
    def __init__(self, excel_file, log_prefix=None, has_page_tags=False, report_focus=None, **kwargs):
        super().__init__(excel_file, log_prefix, has_page_tags, **kwargs)
        self.report_focus = report_focus
        self.get_pages()

    @staticmethod
    def get_tab_groupings(charts):
        tab_color_groupings = []
        chart_hopper = []
        for chart in charts:
            chart_hopper.append(chart)
            if any([
                chart.tab_color is None,
                len(chart_hopper) > 1 and chart.tab_color != chart_hopper[-2].tab_color
            ]):
                current_chart = chart_hopper.pop(len(chart_hopper) - 1)
                if len(chart_hopper) > 0:
                    tab_color_groupings.append(chart_hopper)
                tab_color_groupings.append([current_chart])
                chart_hopper = []
        return tab_color_groupings

    def get_pages(self):
        charts = self.get_report_data()
        for chart_group in self.get_tab_groupings(charts):
            self.pages.append(PPTXPageMeta(chart_group))

    def get_report_data(self):
        for x, y in zip(self.wb.worksheets, self.wb.sheetnames):
            try:
                yield PPTXFlexibleChartMeta(x, y, report_focus=self.report_focus)
            except KeyError:
                print(f'No Data found on {y}')


class DirecTVReportData(ReportData):
    month_variant = 2
    current_time = get_current_time()
    current_year = current_time['year']
    current_month = current_time['month']
    current_month_year = format_date(current_year, current_month)

    def __init__(self, excel_file, log_prefix=None, has_page_tags=False, entertainment=None, **kwargs):
        super().__init__(excel_file, log_prefix, has_page_tags)
        self.overall_df = self.get_report_data()
        self.current_month_str = self.date_manipulator(self.month_variant)
        self.previous_month_str = self.date_manipulator(self.month_variant + 1)
        self.title_months = self.get_month_list()
        self.entertainment = entertainment

        self.get_pages()

    def get_report_data(self):
        overall_df = get_df_from_worksheet(self.excel_file, worksheet=self.wb.sheetnames[0])
        overall_df.columns = get_column_names('dtv')
        overall_df = overall_df.set_index(['Account', 'Competitor No.'])
        return overall_df

    def date_manipulator(self, variant):
        year_number = (self.current_year - 1) if self.current_month < variant else self.current_year
        month_number = self.current_month - variant
        month_number += (13 if month_number < 0 else 12 if month_number == 0 else month_number)
        return format_date(year_number, month_number, 15)

    def get_month_list(self, number_of_months=3, year=False):
        month_deq = deque(month_abbr[1:])
        month_deq.rotate(1 - self.current_month)
        month_lst_full = list(month_deq)[(number_of_months * -1):]
        if year:
            years = [self.current_year for _ in month_lst_full]
            if 'JAN' in month_lst_full:
                jan_idx = month_lst_full.index('JAN')
                years = [x - (1 if idx < jan_idx else 0) for idx, x in enumerate(years)]
            month_lst_full = [f'{x} {y}' for x, y in zip(month_lst_full, years)]
        return month_lst_full

    @staticmethod
    def get_change_label_text(account_df):
        label_values = ["{:.1%}".format(x) for x in account_df['Sales Market Share'].tolist()]
        change_values = [f'({x:.1f})' for x in account_df['Chg Sales Market Share'].tolist()]
        return {idx: [f'{x}{y}'] for idx, (x, y) in enumerate(zip(label_values, change_values))}

    def get_sales_market_share_stacked_bar(self, account_df):
        sms_df = account_df.loc[:, ['Competitor Brand', 'Sales Market Share']].copy()
        sms_df = sms_df.set_index('Competitor Brand')
        sms_df_t = sms_df.transpose()
        chart = PPTXChartMeta(sms_df_t, '100% STACKED BAR')
        chart.label_text = self.get_change_label_text(account_df)
        year_string = f"‘{str(self.current_year)[-2:]} VS. ‘{str(self.current_year - 1)[-2:]}"
        chart.chart_title = f'SALES MARKET SHARE\n{"-".join(self.title_months)} {year_string}'
        chart.chart_style.c_axis_visible = False
        chart.chart_style.legend_location = chart.chart_style.legend_locations.get('bottom')
        chart.decimal_places = 1
        return [chart]

    @staticmethod
    def get_yoy_tables(account_df):
        tables = []
        table_titles = {
            'SALES': ['Competitor Brand', 'YOY Sales Qtr', 'YOY Sales Annual'],
            'TRAFFIC': ['Competitor Brand', 'YOY Traffic Qtr', 'YOY Traffic Annual']
        }
        for title, columns in table_titles.items():
            table_df = account_df[columns]
            table_df = table_df.set_index('Competitor Brand')
            table_instance = PPTXChartMeta(table_df, 'TABLE')
            table_instance.chart_title = f"YEAR-OVER-YEAR {title}"
            table_instance.growth = True
            table_instance.decimal_places = 1
            tables.append(table_instance)
        return tables

    @staticmethod
    def get_change_labels_with_arrows(current_values, change_values):
        data_labels = {x: [] for x in range(0, len(current_values))}

        for series_idx, (current_series, change_series) in enumerate(zip(current_values, change_values)):
            for value_idx, (current_value, change_percent) in enumerate(zip(current_series, change_series)):
                symbol_parameters = {
                    "\u2191": [change_percent > 0],
                    "\u2193": [change_percent < 0],
                    '': [True]
                }
                symbol = get_key_with_matching_parameters(symbol_parameters)
                formatted_current_value = "{:.0%}".format(current_value)
                data_labels[series_idx].append(f'{symbol}\u000A{formatted_current_value}')
        return data_labels

    def get_column_charts(self, account_df):
        column_chart_info = {
            'KEY PERFORMANCE INDICATORS*': {
                'cols': ['Competitor Brand', 'Past Week Trial', 'Redeemed Coupon', 'Ordered LTO', 'Lapsed User'],
                'pp_cols': ['Competitor Brand', 'Chg Past Week Trial', 'Chg Redeemed Coupon', 'Chg Ordered LTO',
                            'Chg Lapsed User'],
                'color': '*GREEN'
            },
            'ORDERING METHOD*': {
                'cols': ['Competitor Brand', 'Order In Store', 'Order Drive thru', 'Order Pickup/Delivery'],
                'pp_cols': ['Competitor Brand', 'Chg Order In Store', 'Chg Order Drive thru',
                            'Chg Order Pickup/Delivery'],
                'color': '*PURPLE'
            }
        }
        charts = []
        for chart in column_chart_info.keys():
            chart_df = trim_df(account_df, column_chart_info[chart]['cols'], 'Competitor Brand')
            change_from_prev_period_df = trim_df(account_df, column_chart_info[chart]['pp_cols'], 'Competitor Brand')
            chart_values_list = chart_df.transpose().values.tolist()
            prev_period_change_list = change_from_prev_period_df.transpose().values.tolist()
            data_labels = self.get_change_labels_with_arrows(chart_values_list, prev_period_change_list)
            chart_instance = PPTXChartMeta(chart_df, 'COLUMN', preferred_color=column_chart_info[chart]['color'])
            chart_instance.chart_title = chart
            chart_instance.label_text = data_labels
            charts.append(chart_instance)
        return charts

    def get_page_1_charts(self, account_df):
        charts = []
        chart_functions = [
            self.get_sales_market_share_stacked_bar,
            self.get_yoy_tables,
            self.get_column_charts
        ]
        for f in chart_functions:
            charts.extend(f(account_df))
        return charts

    def get_page_2_charts(self, account_df):
        line_graphs = {
            'traffic': {
                'cols': ['Competitor Brand', 'TM1', 'TM2', 'TM3', 'TM4', 'TM5', 'TM6', 'TM7'],
                'title': 'TRAFFIC PERFORMANCE\nROLLING THREE-MONTH SYSTEMWIDE'
            },
            'sales': {
                'cols': ['Competitor Brand', 'SM1', 'SM2', 'SM3', 'SM4', 'SM5', 'SM6', 'SM7'],
                'title': 'SALES PERFORMANCE\nROLLING THREE-MONTH SYSTEMWIDE'
            }
        }
        charts = []
        for graph in line_graphs.keys():
            graph_df = trim_df(account_df, line_graphs[graph]['cols'], 'Competitor Brand')
            graph_df.columns = self.get_month_list(7, year=True)
            graph_df = graph_df.transpose()
            graph_instance = PPTXChartMeta(graph_df, 'LINE')
            graph_instance.chart_title = line_graphs[graph]['title']
            graph_instance.decimal_places = 1
            charts.append(graph_instance)
        return charts

    def get_page_3_charts(self, account_df):
        column_list = [
            'Competitor Brand',
            'Overall Ambience/Atmosphere',
            'Video/TV Entertainment',
            'This was the right place for the occasion',
            'Appropriateness for the variety of occasions'
        ]
        change_column_list = [
            'Competitor Brand',
            'Chg Overall Ambience/Atmosphere',
            'Chg Video/TV Entertainment',
            'Chg This was the right place for the occasion',
            'Chg Appropriateness for the variety of occasions'
        ]
        df = trim_df(account_df, column_list, 'Competitor Brand')
        change_from_previous_df = trim_df(account_df, change_column_list, 'Competitor Brand')
        change_from_previous_df.columns = df.columns
        df = df.transpose()
        change_from_previous_df = change_from_previous_df.transpose()
        table_instance = PPTXChartMeta(df, 'TABLE')
        table_instance.change_from_previous = change_from_previous_df
        table_instance.decimal_places = 1
        return [table_instance]

    def get_page_3_body_copy(self):
        body_copy = ['/b/h2Provides Video/TV Entertainment', '/h2Top Box', '']
        segments = ['Quick Service', 'Fast Casual', 'Midscale', 'Casual Dining']
        for segment, value in zip(segments, self.entertainment):
            body_copy.extend([f'/h1{value}', f'/h2{segment}', ''])
        return body_copy

    @staticmethod
    def get_page_3_footer():
        return [
            '/*GREENGreen# = increased from year-end Q2 2020 to year-end Q3 2020',
            '/*ORANGEOrange# = decreased from year-end Q2 2020 to year-end Q3 2020'
        ]

    def get_pages(self):
        page_one_footer = [
            '* Percent shown reflect results from September 2020 YTD',
            f'Arrows indicate change from {self.current_month - 1} YTD to {self.current_month} YTD'
        ]
        account_list = sorted(list(set(self.overall_df.index.get_level_values(0))))
        for account in account_list:
            title = f'{account} Consumer Visit Tracker & Ignite Consumer, {self.current_month_year}'.upper()
            account_df = self.overall_df.loc[account]
            page_1_charts = self.get_page_1_charts(account_df)
            page_1 = PPTXPageMeta(charts=page_1_charts, function='dtv_1', title=title, footer=page_one_footer)
            page_2 = PPTXPageMeta(charts=self.get_page_2_charts(account_df), function='dtv_2')
            if self.has_page_tags:
                page_1.copy = {0: [title]}
                page_2.copy = {0: [title]}
            self.pages.extend([page_1, page_2])
            if all([x != '' for x in self.entertainment]):
                page_3 = PPTXPageMeta(charts=self.get_page_3_charts(account_df), function='dtv_3')
                page_3.copy = {}
                if self.has_page_tags:
                    page_3.copy = {-1: [title]}
                page_3.copy[0] = self.get_page_3_body_copy()
                page_3.footer = self.get_page_3_footer()
                self.pages.append(page_3)


class LTOScorecardReportData(ReportData):
    index_ranges = [[96, 122], [100, 127], [99, 116], [98, 117], [-99999, 99999]]
    lto_scorecard_text = 'templates/import_resources/text_files/lto_scorecard_text.xlsx'

    def __init__(self, excel_file, log_prefix=None, has_page_tags=False, **kwargs):
        super().__init__(excel_file, log_prefix, has_page_tags, **kwargs)
        self.overall_df = self.get_report_data()
        self.company = self.wb.sheetnames[0]
        self.explainer_text_1 = None
        self.explainer_text_2 = None
        self.explainer_text_3 = None

        self.get_explainer_text()
        self.get_pages()

    def get_report_data(self):
        overall_df = get_df_from_worksheet(self.excel_file, worksheet=self.wb.sheetnames[0])
        overall_df.columns = [
            'Concept',
            'Day/Part',
            'Description',
            'PI 2nd Box',
            'PI Top Box',
            'PI Top 2 Box',
            'Purchase Intent',
            'Uniqueness 2nd Box',
            'Uniqueness Top Box',
            'Uniqueness Top 2 Box',
            'Uniqueness',
            'Draw 2nd Box',
            'Draw Top Box',
            'Draw Top 2 Box',
            'Draw',
            'CRV 2nd Box',
            'CRV Top Box',
            'CRV Top 2 Box',
            'Craveability',
            'BF 2nd Box',
            'BF Top Box',
            'BF Top 2 Box',
            'Brand Fit',
            'Anytime',
            'Certain times',
            'A few times',
            'Once',
            'Some Visits',
            'Most Visits',
            'Every Visit',
            'Median willingness to pay',
            'Male',
            'Female',
            'Gen Z',
            'Millennials',
            'Gen X',
            'Baby Boomers',
            '<$45K',
            '$45K-$99K',
            '$100K+',
            'Black/African American',
            'White',
            'Hispanic/Latino'
        ]
        overall_df = overall_df.set_index('Concept')
        for new_column in ['Designation', 'Dine-In', 'Takeout', 'Delivery']:
            overall_df[new_column] = 0
        return overall_df

    def get_explainer_text(self):
        self.explainer_text_1 = get_df_from_worksheet(self.lto_scorecard_text, worksheet=0)
        self.explainer_text_2 = get_df_from_worksheet(self.lto_scorecard_text, worksheet=1)
        self.explainer_text_3 = get_df_from_worksheet(self.lto_scorecard_text, worksheet=2)

    def get_cover_page(self):
        cover = PPTXPageMeta(charts=None, function='ignite cover')
        cover.title = 'Menu Concept Screener Summary'
        cover.copy = {
            0: [f'A review of {self.company} menu concepts'],
            1: [
                'Report Highlights',
                'Methodology',
                'Summary',
                'Benchmarking',
                'Demographics of Potential Purchasers'
            ]
        }
        return [cover]

    def get_methodology_pages(self):
        pages = []
        for explainer in [self.explainer_text_1, self.explainer_text_2, self.explainer_text_3]:
            intro_title = explainer.columns[0]
            text_lst = explainer[intro_title].tolist()
            text_lst.insert(0, intro_title)
            text = {}
            if self.has_page_tags:
                text[-1] = ['Methodology and Definitions']
            text[0] = text_lst
            explainer_page = PPTXPageMeta(charts=None, function='text')
            explainer_page.copy = text
            pages.append(explainer_page)
        return pages

    def get_intro_slide(self):
        """
        The order of text in the dictionaries is based on the placeholder's order of addition to the layout in PPT
        This is a confusing and annoying reality of working with PPT. Text must be added to the dictionary
        in the order that will drop it into PPT placeholders appropriately.
        """
        intro = PPTXPageMeta(charts=None, function='intro')
        intro.copy = {0: ['Recommended Action', 'Details and Qualifiers']}
        if self.has_page_tags:
            intro.copy[-1] = ['Menu Concept Screener Summary']
        intro.copy[1] = ['Proceed Investment', 'N/A', '/h1Proceed Caution', 'N/A', '/h1Stop investment', 'N/A']

        return [intro]

    def get_benchmark_table(self):
        indexes_df = self.overall_df.copy()[['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability', 'Designation']]
        for column in indexes_df.columns:
            column_values = indexes_df[column].tolist()
            indexes_df[column] = [int(x * 100) for x in column_values]

        benchmark_table = PPTXChartMeta(indexes_df, 'TABLE')
        benchmark_table.chart_title = 'CONCEPT BENCHMARKING'
        benchmark_table.indexes = self.index_ranges
        benchmark_table.decimal_places = 0
        benchmark_table.number_format = '0'
        benchmark_table.chart_style.v_banding = True
        benchmark_table.chart_style.h_banding = False
        return [benchmark_table]

    def get_benchmarking_page(self):
        benchmark = PPTXPageMeta(charts=self.get_benchmark_table(), function='table')
        if self.has_page_tags:
            benchmark.copy = {0: ['Concept Scorecards and Benchmarking']}
        benchmark.footer = ['*Index score based on top-box response within daypart-mealpart']
        return [benchmark]

    def get_metrics_top_box(self, concept):
        f1 = self.overall_df.copy().loc[[concept], ['PI 2nd Box', 'PI Top Box', 'PI Top 2 Box']]
        f2 = self.overall_df.copy().loc[[concept], ['Uniqueness 2nd Box', 'Uniqueness Top Box', 'Uniqueness Top 2 Box']]
        f3 = self.overall_df.copy().loc[[concept], ['Draw 2nd Box', 'Draw Top Box', 'Draw Top 2 Box']]
        f4 = self.overall_df.copy().loc[[concept], ['CRV 2nd Box', 'CRV Top Box', 'CRV Top 2 Box']]
        for f in [f1, f2, f3, f4]:
            f.columns = ['2nd Box', 'Top Box', 'Top 2 Box']

        stacked_df = concat_dfs([f1, f2, f3, f4])
        stacked_df.index = ['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability']
        stacked_chart = PPTXChartMeta(stacked_df, 'STACKED COLUMN')
        stacked_chart.top_box = True
        stacked_chart.chart_style.legend_location = stacked_chart.chart_style.legend_locations.get('bottom')
        return stacked_chart

    def get_seasonality_pie(self, concept):
        pie_df = self.overall_df.copy().loc[[concept], ['Anytime', 'Certain times', 'A few times']]
        pie_df = pie_df.transpose()
        pie_chart = PPTXChartMeta(pie_df, 'PIE')
        pie_chart.chart_title = 'Seasonality (Would purchase ______ during the year)'
        return pie_chart

    def get_repeat_trial_pie(self, concept):
        repeat_df = self.overall_df.copy().loc[[concept], ['Once', 'Some Visits', 'Most Visits', 'Every Visit']]
        repeat_df = repeat_df.transpose()
        pie_chart = PPTXChartMeta(repeat_df, 'PIE')
        pie_chart.chart_title = 'Repeat trial'
        return pie_chart

    def get_scorecard_charts(self, concept):
        charts = [
            self.get_metrics_top_box(concept),
            self.get_seasonality_pie(concept),
            self.get_repeat_trial_pie(concept)
        ]
        return charts

    def get_index_shapes(self, concept):
        from_top = 0.75 if self.has_page_tags else 0.46
        index_ranges = {
            'Purchase Intent': [1.22, 0.96],
            'Uniqueness': [1.27, 1.00],
            'Draw': [1.16, 0.99],
            'Craveability': [1.17, 0.98]
        }
        index_title = ParagraphInstance('INDEX', font_color='BLACK', font_size=11, alignment='left')
        shapes = [ShapeMeta(height=0.3, width=9.15, top=from_top, left=3.71, text=index_title, fill_color='GRAY')]

        for idx, column in enumerate(['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability']):
            index_value = self.overall_df.copy().at[concept, column]
            color_parameters = {
                'GREEN': [index_value > max(index_ranges[column])],
                'ORANGE': [index_value < min(index_ranges[column])],
                'BLACK': [True]
            }
            color = get_key_with_matching_parameters(color_parameters)
            text = ParagraphInstance(f'{(index_value * 100):.0f}', font_color=color, font_size=11)
            shapes.append(ShapeMeta(height=0.3, width=0.5, top=from_top, left=[4.7, 6.9, 9.06, 11.27][idx], text=text))
        return shapes

    def get_scorecard_pages(self):
        pages = []
        for concept_idx, concept in enumerate(self.overall_df.index):
            page = PPTXPageMeta(charts=self.get_scorecard_charts(concept), function='lto scorecard')
            page.copy = {}
            if self.has_page_tags:
                page.copy[-1] = ['Concept Scorecards and Benchmarking']
            page.copy[0] = [
                f"${self.overall_df.at[concept, 'Median willingness to pay']:.2f}",
                'Median willingness to pay'
            ]
            page.copy[1] = [
                f'/h2/tag{self.overall_df.iloc[concept_idx, 0].upper()}',
                f'/h1{concept}',
                str(self.overall_df.iloc[concept_idx, 1])
            ]
            page.footer = [
                'Definitions, 2nd Box and Top Box respectively:',
                '/bPurchase Intent#: Likely or very likely to purchase',
                '/bUniqueness#: Unique and very unique',
                '/bDraw#: Likely or much more likely to order from',
                '/bCraveability#: Think I would crave and would definitely crave'
            ]
            page.shapes = self.get_index_shapes(concept)
            pages.append(page)
        return pages

    def get_off_premise_table(self):
        off_premise_df = self.overall_df.copy()[['Dine-In', 'Takeout', 'Delivery']]
        off_premise_table = PPTXChartMeta(off_premise_df, 'TABLE')
        off_premise_table.chart_title = 'OFF-PREMISE POTENTIAL'
        off_premise_table.chart_style.h_banding = True
        return [off_premise_table]

    def get_off_premise_page(self):
        footer = [
            'Q: How would you eat this item? Select all that apply.',
            'Base: potential purchasers (top 2 box purchase intent)'
        ]
        off_premise_page = PPTXPageMeta(charts=self.get_off_premise_table(), function='table', footer=footer)

        if self.has_page_tags:
            off_premise_page.copy = {-1: ['Concept Scorecards and Benchmarking']}
        return [off_premise_page]

    def get_demographics_table(self):
        demographics_df = self.overall_df.copy()[[
            'Male', 'Female', 'Gen Z', 'Millennials', 'Gen X', 'Baby Boomers', '<$45K', '$45K-$99K', '$100K+',
            'Black/African American', 'White', 'Hispanic/Latino'
        ]]
        demographics_df = demographics_df.transpose()
        demographics_df.columns = [str(idx + 1) for idx, _ in enumerate(demographics_df.columns)]
        demographics_table = PPTXChartMeta(demographics_df, 'TABLE')
        demographics_table.highlight = True
        demographics_table.chart_title = 'DEMOGRAPHICS OF POTENTIAL PURCHASERS'
        return [demographics_table]

    def get_demographics_page(self):
        copy = {}
        if self.has_page_tags:
            copy[-1] = ['Concept Scorecards and Benchmarking']
        copy[0] = [f'/h2{idx}. {concept}' for idx, concept in enumerate(self.overall_df.index, 1)]
        footer = ['Potential purchasers=top 2 box purchase intent']
        demographics_page = PPTXPageMeta(self.get_demographics_table(), function='table and text', footer=footer)
        demographics_page.copy = copy
        return [demographics_page]

    def get_end_cap_page(self):
        end_cap_page = PPTXPageMeta(charts=None, function='end cap')
        copy = {}
        if self.has_page_tags:
            copy[-1] = ['Concept Scorecards and Benchmarking']
        copy[0] = ['Alexis Joyce', 'Research Analyst', 'ajoyce@technomic.com']
        copy[1] = ['Mary Clare Metherd', 'Research Analyst', 'mmetherd@technomic.com']
        copy[2] = ['']
        copy[3] = ["So. What's Next?", "Need some more LTO guidance? Reach out to our experts."]
        end_cap_page.copy = copy
        end_cap_page.pictures = [
            'templates/import_resources/headshots/aj.jpg',
            'templates/import_resources/headshots/mcm.jpg'
        ]
        return [end_cap_page]

    def get_pages(self):
        pages = [
            self.get_cover_page(),
            self.get_methodology_pages(),
            self.get_intro_slide(),
            self.get_benchmarking_page(),
            self.get_scorecard_pages(),
            self.get_off_premise_page(),
            self.get_demographics_page(),
            self.get_end_cap_page()
        ]
        for page_list in pages:
            self.pages.extend(page_list)


class ConsumerKPIReportData(ReportData):
    # current_month = 'August 2022'
    # time_period = '(Q1 2021 to Q4 2021)'
    overall_brand_base = '700'
    explainer_text_file = 'templates/import_resources/text_files/consumer_kpi_static_text.xlsx'

    demographic_names = {
        'consumer kpis': {
            'sex': ['Female', 'Male'],
            'generation': ['Generation Z', 'Millennials', 'Generation X', 'Baby Boomers', 'Matures'],
            'ethnicity': ['Female', 'Male', 'Generation Z', 'Millennials', 'Generation X', 'Baby Boomers', 'Matures',
                          'Asian', 'Black/African American', 'Hispanic/Latino', 'Other Ethnicity',
                          'White (non-Hispanic/Latino)', 'Under $25K', '$25K - $50K', '$50K - $75K', '$75K - $100K',
                          '$100K+'],
            'income': ['Under $25K', '$25K - $50K', '$50K - $75K', '$75K - $100K', '$100K+']

        },
        'canada consumer kpis': {
            'sex': ['Female', 'Male'],
            'generation': ['Generation Z', 'Millennials', 'Generation X', 'Baby Boomers', 'Matures'],
            'ethnicity': ['Black/African American', 'Chinese', 'Hispanic/Latino', 'Other Asian', 'Other Ethnicity',
                          'South Asian', 'White (non-Hispanic/Latino)'],
            'income': ['Under $25K', '$25K - $50K', '$50K - $75K', '$75K - $100K', '$100K+']

        },
        'c-store consumer kpis': {
            'sex': ['Female', 'Male'],
            'generation': ['Generation Z', 'Millennials', 'Generation X', 'Baby Boomers', 'Matures'],
            'ethnicity': ['Asian', 'Black/African American', 'White (non-Hispanic/Latino)', 'Hispanic/Latino',
                          'Other Ethnicity'],
            'income': ['Under $25,000', '$25,000-$34,999', '$35,000-$49,999', '$50,000-$74,999', '$75,000-$99,999',
                       '$100,000-$150,000', '$150,000+',]
        }

    }

    def __init__(self, excel_file, log_prefix=None, has_page_tags=False, verbatims=None, report_focus=None,
                 kpi_footer_ranges=[None, None], **kwargs):
        super().__init__(excel_file, log_prefix, has_page_tags, **kwargs)
        self.report_focus = report_focus
        self.overall_df = self.get_data_from_excel(0)
        self.current_month = str(kpi_footer_ranges[0])
        self.time_period = f'({kpi_footer_ranges[1]})'
        self.brand_series = self.overall_df.copy().loc[report_focus]
        self.colloquial_brand = self.overall_df.at[report_focus, 'Brand Name']
        self.alt_visit_names = self.brand_series['VA-1':'VA-6'].values.tolist()
        self.segment = self.get_segment_code()
        self.segment_df = self.get_segment_df()
        self.abbreviation_used = 'AVG' if any([f'{self.segment} AVG' in x for x in self.overall_df.index]) else 'Avg'
        self.importance_df = self.get_attribute_importance_df()
        self.verbatims_dict = verbatims
        self.craveable_verbatims_df = self.get_verbatims_df('overall satisfaction')
        self.satisfaction_verbatims_df = self.get_verbatims_df('craveability')
        self.competitive_df = self.get_competitive_df()
        self.explainer_df = self.get_data_from_excel(0, self.explainer_text_file, set_index=False)
        self.archetypes_df = self.get_data_from_excel(1, self.explainer_text_file)

        self.company = self.wb.sheetnames[0]
        self.get_pages()

    def get_segment_code(self):
        segment = self.brand_series.at['Seg']
        segment = segment.replace('Convenience ', 'C-')
        segment = segment.upper() if all([x not in segment for x in ['C-Store', 'Midscale']]) else segment
        return segment

    def get_segment_df(self):
        index_list = [x for x in self.overall_df.index if x.upper() == f"{self.segment} AVG".upper()]
        return self.overall_df.copy().loc[index_list[0]]

    def get_attribute_importance_df(self):
        df = self.get_data_from_excel(1, set_index=True)
        value_list = df.copy().loc[self.segment].values.tolist()
        df = get_df_from_dict({'Attribute': value_list[0:6], 'Importance': value_list[6:]})
        return df

    def get_competitive_df(self):
        score_comparison_list = [f'{self.segment} {self.abbreviation_used}', self.report_focus]
        for name in self.alt_visit_names:
            if name in self.overall_df.index:
                score_comparison_list.append(name)
            else:
                raise KeyError(name + ' not found in table. Possible Typo or missing data')

        competitive_df = self.overall_df.copy().loc[score_comparison_list]
        for name in competitive_df.index:
            brand_name = competitive_df.at[name, 'Brand Name']
            print(brand_name)
            if not is_null(brand_name):
                competitive_df = competitive_df.rename(index={name: brand_name})
        return competitive_df

    def get_verbatims_df(self, df_name):
        df = self.verbatims_dict[df_name]
        df = df.set_index(df.columns[0])
        df = df.loc[self.colloquial_brand].copy()
        return df

    def get_data_from_excel(self, sheet_idx, excel_file=None, set_index=True):
        excel_file = self.excel_file if excel_file is None else excel_file
        df = get_df_from_worksheet(excel_file, worksheet=sheet_idx)
        if set_index:
            df = df.set_index(df.columns[0])
        return df

    def get_cover_page(self):
        print(self.segment)
        cover = PPTXPageMeta(charts=None, function='ignite cover')
        if self.segment == "C-Store":
            cover.title = f'{self.colloquial_brand} KPI Stats'
            cover.copy = {
                0: ['Top-Line Competitive Brand Assessment'],
                1: [
                    'Brand Health Scorecard',
                    'Overall Satisfaction',
                    'Food & Beverage',
                    'Craveable Items'
                ]
            }
        else:
            cover.title = 'Quarterly Competitive Report'
            cover.copy = {
                0: [f'Created for {self.colloquial_brand}'],
                1: [
                    'Brand Health Scorecard',
                    'Overall Satisfaction',
                    'Food & Beverage',
                    'Intent to Return'
                ]
            }
        shape_text = ParagraphInstance(self.current_month.upper(), font_color='black', font_size=10, alignment='center')
        cover.shapes = [ShapeMeta(height=2.19, width=0.4, top=10.65, left=5.0, text=shape_text, fill_color='white')]
        return [cover]

    def get_demographic_column_names(self):
        parameters = {
            "c-store consumer kpis": ["C-Store" in self.segment],
            "canada consumer kpis": ["Chinese" in self.overall_df.columns],
            "consumer kpis": [True]
        }
        return self.demographic_names[get_key_with_matching_parameters(parameters)]

    def get_demographic_skew_intro(self):
        archetype = self.archetypes_df.loc[self.brand_series.at["EaterArchetype"], "Name"]
        copy = {0: [
            f'{self.colloquial_brand} guest Eater Archetype skew: {archetype}',
            str(self.archetypes_df.loc[self.brand_series.at["EaterArchetype"], "Description"])
        ]}
        if self.has_page_tags:
            copy[-1] = [f"COMPETITIVE BRAND PERFORMANCE | {self.colloquial_brand}"]
        copy[1] = ['About Consumer Tracking'] + self.explainer_df[self.explainer_df.columns[0]].tolist()
        copy[2], copy[3], copy[4] = [], [], [' ']

        copy_keys = [2, 2, 3, 3]
        demographics = self.get_demographic_column_names()
        for idx, (category, demographic_list) in enumerate(demographics.items()):
            demographics_frame = self.brand_series.copy()[demographic_list].transpose()
            demographics_values = [demographics_frame.at[x] for x in demographic_list]
            segment_values = [self.segment_df.at[x] for x in demographic_list]
            segment_skews = [x - y for x, y in zip (demographics_values, segment_values)]

            max_skew = max(segment_skews)
            max_skew_index = segment_skews.index(max_skew)
            max_skew_value = demographics_values[max_skew_index]
            max_skew_brand_demographic = demographic_list[max_skew_index]
            max_skew_segment_value = self.segment_df.at[max_skew_brand_demographic]

            segment_max_val = f'/b{t_round(max_skew_segment_value, 3):.1%}#'

            copy[copy_keys[idx]].append(f'/h1{t_round(max_skew_value, 3):.1%}')

            income_phrase = '/clearof frequent guests have a household income of #'
            phrase = income_phrase if idx == 3 else '/clearof frequent guests are#'
            text = f'{phrase} {max_skew_brand_demographic} /clearcompared to# {segment_max_val} across the {self.segment} segment'
            copy[copy_keys[idx]].append(text)

        base_size = int(self.brand_series["Base Size"])
        footer = [
            f'{self.segment} Base: {base_size} once a month+ {self.segment} consumer, {self.time_period}',
            'Source: Ignite Consumer'
        ]

        page = PPTXPageMeta(charts=None, function='intro', footer=footer)
        page.copy = copy
        return [page]

    def get_visit_alternatives_copy(self):
        footer = [
            f'Base: {int(self.brand_series["Base Size"])} recent {self.colloquial_brand} guests, {self.time_period}',
            'Source: Ignite Consumer'
        ]
        visit_type = 'Visit a c-store or restaurant' if 'C-Store' in self.segment else "Visit a restaurant"
        copy = {}
        if self.has_page_tags:
            copy[-1] = [f"COMPETITIVE BRAND PERFORMANCE | {self.colloquial_brand}"]
        copy[0] = [
            f'{t_round(self.brand_series[visit_type], 3): .1%}'.strip(),
            f'would have gone to another would have gone to another C-store or restaurant as an alternative to {self.colloquial_brand}'
        ]
        return footer, copy

    def get_visit_alternatives_chart(self):
        values = self.brand_series['VA1-score':'VA6-score']
        values.index = [self.overall_df.at[x, 'Brand Name'] for x in self.alt_visit_names]
        alt_visit_values_df = values.to_frame()
        chart = PPTXChartMeta(alt_visit_values_df, 'COLUMN')
        chart.chart_title = 'Percent of Consumers who Considered Visiting _________'.upper()
        chart.decimal_places = 1
        return [chart]

    def get_visit_alternatives(self):
        page = PPTXPageMeta(charts=self.get_visit_alternatives_chart(), function='chart and text')
        page.footer, page.copy = self.get_visit_alternatives_copy()
        return [page]

    def find_attribute_variation(self, attribute):
        """
        This breaks apart the attribute name and looks for alternative name that shares the
        most similar vocabulary.
        """
        attribute_split = attribute.split()
        columns = self.competitive_df.columns.tolist()
        match_count = [[x in column for x in attribute_split].count(True) for column in columns]
        max_match_index = match_count.index(max(match_count))
        max_match_column = columns[max_match_index]
        return max_match_column

    def get_attribute_importance_charts(self):
        charts = []
        maximum_list = []
        for attribute in self.importance_df['Attribute'].tolist():
            attribute_title = self.find_attribute_variation(attribute)
            attribute_df = self.competitive_df.copy()[[attribute_title]]
            attribute_df = attribute_df.sort_values(by=attribute_title, ascending=False)
            attribute_df = attribute_df.fillna(0.0)
            maximum_list.append(attribute_df.iat[0, 0])
            chart = PPTXChartMeta(attribute_df, 'BAR')
            chart.chart_title = attribute_title.upper()
            chart.decimal_places = 1
            chart.highlight = True
            emphasis = f"{self.segment} Avg"
            chart.emphasis.append(emphasis)
            chart.de_emphasis.extend([x for x in attribute_df.index if x not in [emphasis, self.colloquial_brand]])
            charts.append(chart)
        for chart in charts:
            chart.chart_style.v_axis_maximum = max(maximum_list) + 0.2
            chart.chart_style.v_axis_minimum = 0
        return charts

    def get_attribute_importance_copy(self):
        footer = [
            'Q: Based on your recent visit, how would you rate the chain on the following?',
            f"Base: {self.overall_brand_base} recent guests per brand {self.time_period}"
            'Showing percentage selecting very good (top-box rating)',
            'Source: Ignite Consumer'
        ]
        copy = {}
        if self.has_page_tags:
            copy[-1] = [f"COMPETITIVE BRAND PERFORMANCE | {self.colloquial_brand}"]
        copy[0] = [f'/tagTop six visit factors when selecting a {self.segment} for a meal']
        importance, attribute = self.importance_df['Importance'].tolist(), self.importance_df['Attribute'].tolist()
        segment_attribute_list = [[f"/b{t_round(value, 3): .1%}# {name}"] for value, name in zip(importance, attribute)]
        for lst in segment_attribute_list:
            copy[0].extend(lst)
        return footer, copy

    def get_attribute_importance(self):
        page = PPTXPageMeta(charts=self.get_attribute_importance_charts(), function='TT_6_Chart_&_Text')
        page.footer, page.copy = self.get_attribute_importance_copy()
        return [page]

    def get_most_craveable_items_footer(self):
        base_number = f'{int(self.brand_series["Craveable Base"]):,}'
        footer = [
            f' Base: {base_number} recent {self.colloquial_brand} guests {self.time_period}',
            'Source: Ignite Consumer'
        ]
        return footer

    def get_craveable_verbatims(self):
        positives = {'crave', 'i love t', 'excellent '}
        craveable = []
        try:
            craveable = self.craveable_verbatims_df[self.craveable_verbatims_df.columns[1]]
        except AttributeError:
            pass
        verbatims = [x for x in craveable if any([y.lower() in str(x).lower() for y in positives])]
        copy = {}
        if self.has_page_tags:
            copy[-1] = [f"COMPETITIVE BRAND PERFORMANCE | {self.colloquial_brand}"]
        if len(verbatims) > 0:
            copy[0] = [f"/h5{x}" for x in verbatims[:10]]
        return copy

    def get_most_craveable_items_chart(self):
        categories = [x for x in self.brand_series['craveable_item1':'craveable_item5'].tolist()]
        craveable_values = self.brand_series['score1':f'score{len(categories)}']
        craveable_values.index = categories
        craveable_values_df = craveable_values.to_frame()
        craveable_values_df = craveable_values_df.dropna()
        chart = PPTXChartMeta(craveable_values_df, 'COLUMN')
        chart.chart_title = 'Most Craveable Items'.upper()
        chart.decimal_places = 1
        return [chart]

    def get_most_craveable_items(self):
        chart = self.get_most_craveable_items_chart()

        copy = self.get_craveable_verbatims()
        try:
            len(copy[0])
            page_function = 'chart and text'
        except KeyError:
            page_function = 'chart'

        page = PPTXPageMeta(charts=chart, function=page_function)
        page.copy = copy
        page.footer = self.get_most_craveable_items_footer()
        return [page]

    def get_overall_satisfaction_footer(self):
        footer = [
            'Q: Based on your recent visit, how would you rate the chain on the following?',
            f"Total base: {self.overall_brand_base} recent guests per brand {self.time_period}",
            'Showing percentage selecting very good (top-box rating)'
        ]
        return footer

    def get_satistfaction_verbatims(self):
        positives = {'crave', 'i love t', 'excellent ', ' enjoy', 'best around', 'has the best', 'good quality'}
        craveable = []
        try:
            craveable = self.craveable_verbatims_df[self.craveable_verbatims_df.columns[1]]
        except AttributeError:
            pass
        verbatims = [x for x in craveable if any([y.lower() in str(x).lower() for y in positives])]
        copy = {}
        if self.has_page_tags:
            copy[-1] = [f"COMPETITIVE BRAND PERFORMANCE | {self.colloquial_brand}"]
        if len(verbatims) > 0:
            copy[0] = [f"/h5{x}" for x in verbatims[:10]]
        return copy

    def get_overall_satisfaction_chart(self):
        overall_score_df = self.competitive_df.copy()[['Overall Rating']]
        overall_score_df = overall_score_df.sort_values(by='Overall Rating', ascending=False)
        chart = PPTXChartMeta(overall_score_df, 'BAR')
        chart.chart_title = 'Overall Visit Satisfaction'.upper()
        chart.decimal_places = 1
        chart.highlight = True
        emphasis = f"{self.segment} Avg"
        chart.emphasis.append(emphasis)
        chart.de_emphasis.extend([x for x in overall_score_df.index if x not in [emphasis, self.colloquial_brand]])
        return [chart]

    def get_overall_satisfaction(self):
        footer = self.get_overall_satisfaction_footer()
        chart = self.get_overall_satisfaction_chart()

        copy = self.get_satistfaction_verbatims()
        try:
            len(copy[0])
            page_function = 'chart and text'
        except KeyError:
            page_function = 'chart'
        page = PPTXPageMeta(charts=chart, function=page_function, footer=footer)
        page.copy = copy
        return [page]

    def get_end_cap_page(self):
        end_cap_page = PPTXPageMeta(charts=None, function='end cap')
        end_cap_page.copy = {}
        if self.has_page_tags:
            end_cap_page.copy[-1] = [f"COMPETITIVE BRAND PERFORMANCE | {self.colloquial_brand}"]

        end_cap_page.copy[0] = ['Robert Byrne', 'Director, Research and Insights' 'rbyrne@technomic.com']
        end_cap_page.copy[1] = ['Britany Trujillo', 'Manager, Research & Insights', 'btrujillo@technomic.com']
        end_cap_page.copy[2] = [' ']
        end_cap_page.copy[3] = ["So. What's Next?", "Need some consumer questions answered? Reach out to our experts."]

        end_cap_page.pictures = [
            'templates/import_resources/headshots/rb.jpg',
            'templates/import_resources/headshots/bt.jpg',
            'templates/import_resources/headshots/ignite.png'
        ]
        return [end_cap_page]

    def get_pages(self):
        pages = [
            self.get_cover_page(),
            self.get_demographic_skew_intro(),
            self.get_visit_alternatives(),
            self.get_attribute_importance(),
            self.get_most_craveable_items(),
            self.get_overall_satisfaction(),
            self.get_end_cap_page()
        ]
        for page_list in pages:
            self.pages.extend(page_list)
