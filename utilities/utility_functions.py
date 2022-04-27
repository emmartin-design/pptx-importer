import decimal
import pandas as pd
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

country_list = [
    'Argentina',
    'Australia',
    'Brazil',
    'Canada',
    'Chile',
    'China',
    'Colombia',
    'France',
    'Germany',
    'India',
    'Indonesia',
    'Japan',
    'Malaysia',
    'Mexico',
    'Philippines',
    'Russia',
    'Saudi Arabia',
    'Singapore',
    'South Africa',
    'South Korea',
    'Spain',
    'Thailand',
    'UAE',
    'U.K.',
    'U.S.'
]


def combine_lists(list_of_lists: list):
    new_list = []
    for lst in list_of_lists:
        new_list.extend(lst)
    return new_list


def concat_dfs(list_of_dfs):
    return pd.concat(list_of_dfs)


def format_as_float(value):
    try:
        return float(value)
    except ValueError:
        return 0.0


def format_date(year_number, month_number, day=15, string_format='%B %Y'):
    return datetime(year_number, month_number, day).strftime(string_format)


def get_df_from_dict(data_dict):
    return pd.DataFrame.from_dict(data_dict)


def get_df_from_worksheet(path, worksheet):
    return pd.read_excel(path, sheet_name=worksheet)


def get_current_time():
    time = {
        'year': datetime.now().year,
        'month': datetime.now().month,
    }
    return time


def get_key_with_matching_parameters(parameter_dict):
    """
    Used for case handling. If any of the statements in the dict values lists for each key are true,
    the key will be returned as the selected item.
    Only pass values with boolean statements.
    """
    for key, params in parameter_dict.items():
        if any(params):
            return key


def is_null(value):
    return pd.isnull(value)


def pivot_dataframe(df, pivot_column):
    try:
        df = df.pivot(columns=pivot_column)
        df = df.droplevel(0, axis='columns')
    except KeyError:
        pass
    return df


def replace_chars(original_string, *args):
    """
    Submit list of tuples for args, with original character and replacement character
    e.g., (' ', '_')
    """
    new_string = str(original_string)
    for arg in args:
        new_string = new_string.replace(arg[0], arg[1], -1)
    return new_string


def trim_df(df, desired_columns: list, new_index=None):
    new_df = df[desired_columns]
    if new_index is not None:
        new_df = new_df.set_index(new_index)
    return new_df


def t_round(value, decimal_places: int = 0):
    """
    Python rounds thus: round to nearest, ties to even
    Technomic rounds thus: round half up (school rounding)
    """
    decimal.getcontext().rounding = decimal.ROUND_HALF_UP
    rounded = round(decimal.Decimal(str(value)), decimal_places)
    if decimal_places == 0:
        if is_null(value):
            return 0
        return int(float(rounded))
    return float(rounded)


def get_column_names(report_type):
    name_options = {
        'dtv': [
            'Account',
            'Competitor No.',
            'Competitor Brand',
            'Sales Market Share',
            'YOY Sales Qtr',
            'YOY Sales Annual',
            'YOY Traffic Qtr',
            'YOY Traffic Annual',
            'Past Week Trial',
            'Redeemed Coupon',
            'Ordered LTO',
            'Lapsed User',
            'Order In Store',
            'Order Drive thru',
            'Order Pickup/Delivery',
            'Provides TV/Entertainment',
            'Overall Ambience/Atmosphere',
            'Video/TV Entertainment',
            'This was the right place for the occasion',
            'Appropriateness for the variety of occasions',
            'Chg Sales Market Share',
            'Chg YOY Sales Qtr',
            'Chg YOY Sales Annual',
            'Chg YOY Traffic Qtr',
            'Chg YOY Traffic Annual',
            'Chg Past Week Trial',
            'Chg Redeemed Coupon',
            'Chg Ordered LTO',
            'Chg Lapsed User',
            'Chg Order In Store',
            'Chg Order Drive thru',
            'Chg Order Pickup/Delivery',
            'Chg Provides TV/Entertainment',
            'Chg Overall Ambience/Atmosphere',
            'Chg Video/TV Entertainment',
            'Chg This was the right place for the occasion',
            'Chg Appropriateness for the variety of occasions',
            'Sales Column PH',
            'SM1',
            'SM2',
            'SM3',
            'SM4',
            'SM5',
            'SM6',
            'SM7',
            'Traffic Column PH',
            'TM1',
            'TM2',
            'TM3',
            'TM4',
            'TM5',
            'TM6',
            'TM7',
        ]
    }
    return name_options.get(report_type)
