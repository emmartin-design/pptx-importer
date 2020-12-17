from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches
from pptx.enum.chart import XL_LEGEND_POSITION
from datetime import datetime
import os
from calendar import month_abbr
from collections import deque
import logging

# APP CONFIGUATION ####################################################################################################

# Version variable for future automatic updates
version = '1.0.1 Beta'

# Default Template
default_template = 'templates/DATAIMPORT.pptx'

# Creates a new log for each day
logging_file_name = 'gen_' + datetime.now().strftime('%y%B%d') + '.log'
log_dir = os.path.join(os.path.normpath(os.getcwd()), 'logs')
log_fn = os.path.join(log_dir, logging_file_name)
logging.basicConfig(filename=log_fn, level=logging.DEBUG, format='%(lineno)d:%(levelname)s:%(message)s')


# Handle GUI updates and logging
def log_entry(msg, level='info', app_holder=None, fieldno=None):
    color_selector = {'warning': 'yellow', 'info': 'light gray'}

    if level == 'warning':
        logging.warning(msg)
    else:
        logging.info(msg)

    if app_holder is not None:
        app_holder.setStatusbar(msg, field=fieldno)
        app_holder.setStatusbarBg(color_selector[level], field=fieldno)
        app_holder.topLevel.update()


# FILE CONFIGURATION ##################################################################################################

# Defines available import types 'Consumer KPIs', 'Menu Scorecards' to be added
report_config = {
    'General Import': {'report type': 'general', 'report list': [None], 'report suffix': ''},
    'Global Navigator Country Reports': {'report type': 'global',
                                         'report list': ['Argentina', 'Australia', 'Brazil', 'Canada', 'Chile', 'China',
                                                         'Colombia', 'France', 'Germany', 'India', 'Indonesia', 'Japan',
                                                         'Malaysia', 'Mexico', 'Philippines', 'Russia', 'Saudi Arabia',
                                                         'Singapore', 'South Africa', 'South Korea', 'Spain',
                                                         'Thailand', 'United Arab Emirates', 'United Kingdom',
                                                         'United States'],
                                         'report suffix': 'Country Report'
                                         },
    'LTO Scorecard Report': {
        'report type': 'lto',
        'report list': [None],
        'report suffix': '',
        'type':'lto',
        'report title': 'Menu Concept Screener Summary',
        'scorecard chart count': 4,
        'split demographics': False,
        'table of contents': ['Report Highlights', 'Methodology', 'Summary', 'Benchmarking',
                              'Demographics of Potential Purchasers'],
        'index column names': ['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability', 'Designation'],
        'off-premise potential names': ['Dine-In', 'Takeout', 'Delivery'],
        'index ranges': [[96, 122], [100, 127], [99, 116], [98, 117]],
        'additional files': 'templates/import_resources/text_files/lto_scorecard_text.xlsx'
    },
    'Value Scorecard Report': {
        'report type': 'value',
        'report list': [None],
        'report suffix': '',
        'type':'value',
        'report title': 'Value Concept Screener Summary',
        'scorecard chart count': 6,
        'split demographics': False,
        'table of contents': ['Report Highlights', 'Methodology', 'Summary', 'Benchmarking',
                              'Demographics of Potential Purchasers'],
        'index column names': ['Purchase Intent Percentile', 'Value Percentile', 'Draw Percentile',
                               'Craveability Percentile', 'Quality Percentile'],
        'index ranges': None,
        'additional files': 'templates/import_resources/text_files/valuelto_scorecard_text.xlsx'
    },
    'DirecTV Scorecard': {'report type': 'dtv', 'report list': [None], 'report suffix': ''}
}

# Slide functions for static data report imports
slide_function_options = {
        'cover': {'body': 2, 'title': 1, 'chart': 0, 'table': 0, 'picture': 0},
        'intro': {'body': 6, 'title': 0, 'chart': 0, 'table': 0, 'picture': 0},
        'text': {'body': 2, 'title': 0, 'chart': 0, 'table': 0, 'picture': 0},
        'table and text': {'body': 2, 'title': 0, 'chart': 0, 'table': 1, 'picture': 0},
        'table': {'body': 1, 'title': 0, 'chart': 0, 'table': 1, 'picture': 0},
        'full page chart': {'body': 1, 'title': 0, 'chart': 1, 'table': 0, 'picture': 0},
        'end wrapper': {'body': 5, 'title': 0, 'chart': 0, 'table': 0, 'picture': 3},
        'two charts, no text': {'body': 1, 'title': 0, 'chart': 2, 'table': 0, 'picture': 0},
        'value scorecard': {'body': 2, 'title': 0, 'chart': 6, 'table': 0, 'picture': 0},
        'dtv': {'body': 1, 'title': 0, 'chart': 3, 'table': 2, 'picture': 0},
    }

# Separate from logging, places error notes in the deck
error_dict = {
    '100': ['ERROR: Data Selection Error. Chart Import Aborted', 'high'],
    '103': ['ERROR: No Data Question selected.', 'med'],
    '104': ['ERROR: No Data Base selected. Please select', 'med'],
    '105': ['ERROR: Data cannot be sorted for data imports of this configuration', 'med'],
    '106': ['NOTE: Blank Cells Selected in original data.', 'med'],
    '107': ['NOTE: Formulas detected in cells. Data not read.', 'high'],
    '404': ['ERROR: Chart type not recognized.', 'high'],
    '405': ['NOTE: Chart type was automatically selected.', 'high'],
    '501': ['ERROR: No data selected in worksheet', 'high'],
    '801': ['ERROR: Chart placeholder type not found. Import aborted.', 'high']
}


# PAGE CONFIGURATION ##################################################################################################

def assign_page_config(chart_count=0, table_count=0, has_stat=False, title=None, tag=None, function=None,
                       full_chart=False, page_copy=None):
    if page_copy is None:
        page_copy = []

    page_config = {
        'number of charts': chart_count,
        'number of tables': table_count,
        'has stat': has_stat,
        'page title': title,
        'section tag': tag,
        'page copy': page_copy,
        'footer': [],
        'page img': [],
        'function': function,
        'full page chart': full_chart,
        'callouts': {}
    }

    return page_config


def assign_layout_config(name=None):

    layout_config = {
        'name': name,
        'chart count': 0,
        'table count': 0,
        'body count': 0,
        'title count': 0,
        'picture count': 0,
        'width test': [],
        'preferred': False
    }

    return layout_config


# CHART CONFIGUATION ##################################################################################################

# Used to determine what kind of information goes into placeholder
placeholder_types = [
    'CHART',
    'TABLE',
    'SLIDE_NUMBER',
    'FOOTER',
    'TITLE',
    'BODY',
    'PICTURE'
]

# Used to assign charts to placeholders
chart_types = {
    'COLUMN': XL_CHART_TYPE.COLUMN_CLUSTERED,
    'STACKED COLUMN': XL_CHART_TYPE.COLUMN_STACKED,
    '100% STACKED COLUMN': XL_CHART_TYPE.COLUMN_STACKED,
    'PIE': XL_CHART_TYPE.PIE,
    'BAR': XL_CHART_TYPE.BAR_CLUSTERED,
    'STACKED BAR': XL_CHART_TYPE.BAR_STACKED,
    '100% STACKED BAR': XL_CHART_TYPE.BAR_STACKED,
    'LINE': XL_CHART_TYPE.LINE,
    'TABLE': None,
    'STAT': None,
    'PICTURE': None
}

# Legend locations based on chart types
legend_locations = {
    'bottom': XL_LEGEND_POSITION.BOTTOM,
    'right': XL_LEGEND_POSITION.RIGHT
}

# Assigns monochromatic color styles (outdated in PPT, but still functional)
chart_styles = {
    'blueberry': 3,
    'strawberry': 4,
    'mint': 5,
    'mandarin': 6,
    'sultana': 7,
    'grape': 8
}

# Used to assign colors to all objects
brand_colors = {
    'default': MSO_THEME_COLOR.TEXT_1,
    'blueberry': MSO_THEME_COLOR.ACCENT_1,
    'strawberry': MSO_THEME_COLOR.ACCENT_2,
    'mint': MSO_THEME_COLOR.ACCENT_3,
    'mandarin': MSO_THEME_COLOR.ACCENT_4,
    'sultana': MSO_THEME_COLOR.ACCENT_5,
    'grape': MSO_THEME_COLOR.ACCENT_6
}

# Used to change colors of chart points based on list contents
chart_color_variations = {
    'emphasis': ['OVERALL', 'SUM', 'TOTAL'],
    'de-emphasis': ['OTHER', 'SAME', 'PREFER NOT TO SAY', 'NEVER', 'OTHER: PLEASE SPECIFY'],
    'highlight': []
}

# Dummy data to insert when major data errors arise.
data_error = {
    'categories': ['data mismatch', 'error', 'data mismatch', 'error'],
    'error': [0, 0, 0, 0]
}


# Returns new dictionary values each time.
def assign_chart_config(
        intended_chart='BAR', top_box=False, heatmap=False, highlight=False, indexes=None, preferred_color='blueberry',
        banding='rows', title=None, legend_location='DEFAULT', category_axis=True, max_values=None, growth=False,
        pct_dec_places=0, force_int=False, label_txt=None, has_data=False, tab_color=None,
        state='visible', dq=None, note=None, data_labels=True):

    # Assigning empty lists here to Avoid bugs
    if max_values is None:
        max_values = []

    if dq is None:
        dq = []

    if note is None:
        note = []

    chart_config = {
        # Data handling
        '*TRANSPOSE': False,
        '*SORT': False,
        'vertical series': False,
        # Chart data formatting
        '*FORCE PERCENT': False,
        '*MEAN': False,
        '*FORCE FLOAT': False,
        '*FORCE CURRENCY': False,
        '*FORCE INT': force_int,
        '*TOP 5': False,
        '*TOP 10': False,
        '*TOP 20': False,
        'number format': '0%',
        'dec places': pct_dec_places,
        # Chart Styling
        '*TOP BOX': top_box,
        '*HEAT MAP': heatmap,
        '*HIGHLIGHT': highlight,
        'intended chart': intended_chart,
        'preferred color': preferred_color,
        'banding': banding,
        'cat axis': category_axis,
        'chart title': title,
        'legend loc': legend_location,
        'max values': max_values,
        'data labels': data_labels,
        # Page Styling
        '*INDEXES': indexes,
        '*GROWTH': growth,
        # Formatting checks and preflights
        'percent check': True,
        'error list': [],
        'notes': [],
        'chart chosen': False,
        'directional check': False,
        # Text Additions
        'title question': [],
        'bases': [],
        'data question': dq,
        'copy': [],
        'label text': label_txt,
        'note': note,
        # Data reading/cleanup
        'has data': has_data,
        'tab color': tab_color,
        'sheet state': state
    }

    if True in [chart_config['*MEAN'], chart_config['*FORCE FLOAT'],
                chart_config['*FORCE CURRENCY'], chart_config['*FORCE INT']]:
        chart_config['percent check'] = False

    return chart_config


# TEXT FORMATTING #####################################################################################################

# Exceptions for cell values
exception_list = ['-', '`', '*', '**', '[', ']', 'u', 'N/A', 'NA']
exception_letter_string = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

# List of words to remove from titles
stopwords = {'PLEASE', 'FOODSERVICE', 'THE', 'WHAT', 'WHO', 'HOW', 'WHICH', 'WHEN', 'IS', 'ARE', 'CHECK', 'CHOOSE',
             'SELECT', 'OPTION', 'ONLY', 'MAIN', 'ONE', 'YOU', 'YOUR', 'MANY', 'DO', 'DOES', 'OF…', 'IF', 'THE',
             'ON', 'THAT', 'ALL', 'PART', 'FOLLOWING', 'BEST', 'DESCRIBES', 'APPLY', 'MARKET', 'OPERATION',
             'REPRESENT', 'THINKING', 'BACK', 'GENERALLY', 'SPEAKING', 'RESTAURANT', 'SERVES'
             }


def month_lst(number_of_months, last_month, year=None, separator='\n'):
    # Creates a list of months equal to length of provided number of months, ending with the last month
    # Adds years as needed
    month = last_month + 1
    month_deq = deque(month_abbr[1:])
    month_deq.rotate(1 - month)
    month_lst_full = list(month_deq)

    if year is not None:
        year -= 1
        for m_idx, m in enumerate(month_lst_full):
            if 'JAN' in str(m).upper():
                year += 1
            month_lst_full[m_idx] = str(m) + separator + str(year)

    return month_lst_full[(number_of_months * -1):]


def title_cleaner(title):
    # splits string into list of words and removes punctuation
    txt = title
    for puncuation in ['?', '.', ',']:
        txt = txt.replace(puncuation, '')
    title_words = txt.upper().split()
    print(title_words)
    result_words = [word for word in title_words if word not in stopwords]
    try:
        if result_words[-1] in ['IN', 'OF']:
            result_words = result_words[:-1]
        if result_words[0] in ['IN', 'OF']:
            result_words = result_words[1:]
    except IndexError:
        pass

    # If title is empty, returns original (e.g., "Which One" to "")
    if len(result_words) == 0:
        new_title = title.upper()
    else:
        # Joins list back to string
        new_title = ' '.join(result_words)
    print(new_title)
    return new_title


# REPORT-SPECIFIC CONFIGURATION #######################################################################################

idx_roundup_columns = {
    'lto': [
        'Concept',
        'Purchase Intent',
        'Uniqueness',
        'Draw',
        'Craveability',
        'Designation'
    ],
    'value': [
        'Concept',
        'Purchase Intent Percentile',
        'Value Percentile',
        'Draw Percentile',
        'Craveability Percentile',
        'Quality Percentile'
    ]
}

scorecard_footercopy = {
    'lto': [
        ' '
    ],
    'value': [
        'Definitions, 2nd Box and Top Box respectively:',
        'Value: Good or Great value',
        'Purchase Intent: Likely or Very Likely to Purchase',
        'Draw: Likely or much more likely to order from',
        'Craveability: Craveable or extremely craveable',
        'Quality: Good or High Quality'
    ]
}

# Used to slim possilbe layouts used for static-data report types
preferred_layouts = {
    'lto': [
        'IGNITE_TitlePage',
        'TT—Subsection Intro, Main Ideas',
        'TT—Full Text',
        'TT_Primary_Table',
        'TT_Primary_Chart',
        'TT_6_Chart_&_Text',
        'TT_3_Chart_Dashboard_Flipped',
        'TT_Primary_Table_&_Text',
        'TT_End Wrapper_w_Photos'
    ],
    'value': [
        'IGNITE_TitlePage',
        'TT—Subsection Intro, Main Ideas',
        'TT—Full Text',
        'TT_Primary_Table',
        'TT_Primary_Chart',
        'TT_6_Chart_&_Text',
        'TT_3_Chart_Dashboard_Flipped',
        'TT_Primary_Table_&_Text',
        'TT_End Wrapper_w_Photos'
    ],
    'dtv': [
        'TT_3_chart_2_table',
        'TT_Two_Chart_Equal',
        'TT_Primary_Table_&_Text'
    ]
}

valcover = [
    'Report Highlights',
    'Methodology',
    'Summary',
    'Benchmarking',
    'Scorecards',
    'Demographics of Potential Purchasers'
]

index_loc = [
    Inches(4.85),
    Inches(8.02),
    Inches(10.45),
    Inches(11.327),
    Inches(12.07)
]

scorecard_dfcols = {
    'lto': [
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
    ],

    'value': [
        'Company',
        'Concept',
        'Description',
        'PI 2nd Box',
        'PI Top Box',
        'PI Top 2 Box',
        'Purchase Intent Percentile',
        'Val 2nd Box',
        'Val Top Box',
        'Val Top 2 Box',
        'Value Percentile',
        'Draw 2nd Box',
        'Draw Top Box',
        'Draw Top 2 Box',
        'Draw Percentile',
        'CRV 2nd Box',
        'CRV Top Box',
        'CRV Top 2 Box',
        'Craveability Percentile',
        'Quality 2nd Box',
        'Quality Top Box',
        'Quality Top 2 Box',
        'Quality Percentile',
        'Only the offer',
        'Additional foods and/or beverages',
        'Unsure',
        'Portion',
        'Taste',
        'Quality',
        'Fits Budget',
        'Once',
        'Some visits',
        'Most visits',
        'Every visit',
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
    ],

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

demographic_cats = ['Democat', 'Gender', 'Gender',
                    'Generation', 'Generation', 'Generation', 'Generation', 'Generation',
                    'Household Income', 'Household Income', 'Household Income', 'Household Income', 'Household Income',
                    'Household Income', 'Ethnicity', 'Ethnicity', 'Ethnicity', 'Ethnicity'
                    ]

mock_chart = {
    'categories': ['Item 1', 'Item 2', 'Item 3', 'Item 4'],
    'Fits somewhat': [0.31, 0.31, 0.32, .31],
    'Is a perfect fit': [0.19, 0.28, 0.34, 0.19],
    'Top 2 box': [0.5, 0.59, 0.66, 0.5]
}

# COMPANY-SPECIFIC VARIABLES #########################################################################################

employees = {
    'lh': ['Lauren Hallow', 'Senior Manager, Research and Insights', 'lhallow@technomic.com'],
    'jc': ['Jenna Carroll', 'Research Analyst', 'jrcarroll@technomic.com']
}

# ROUNDING MODULE ####################################################################################################

def trueround(number, places=0):
    # To add options for rounding types
    place = 10**(places)
    rounded = (int(number*place + 0.5 if number>=0 else -0.5))/place
    if rounded == int(rounded):
        rounded = int(rounded)
    return rounded