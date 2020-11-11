from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches
from pptx.enum.chart import XL_LEGEND_POSITION
from calendar import month_abbr
from collections import deque
import logging

import re
import string


# Version variable for future automatic updates
version = '1.0.0 Beta'



# Default Template
defaulttemplate = 'templates/DATAIMPORT.pptx'

# Defines available import types 'Consumer KPIs', 'Menu Scorecards' to be added
reporttypes = [
    'General Import',
    'Global Navigator Country Reports',
    'LTO Scorecard Report',
    'Value Scorecard Report',
    'DirecTV Scorecard'
]

phtypes = [
    'CHART',
    'TABLE',
    'SLIDE_NUMBER',
    'FOOTER',
    'TITLE',
    'BODY',
    'PICTURE'
]

charttypelist = {
    'COLUMN': XL_CHART_TYPE.COLUMN_CLUSTERED,
    'STACKED COLUMN': XL_CHART_TYPE.COLUMN_STACKED,
    '100% STACKED COLUMN': XL_CHART_TYPE.COLUMN_STACKED,
    'PIE': XL_CHART_TYPE.PIE,
    'BAR': XL_CHART_TYPE.BAR_CLUSTERED,
    'STACKED BAR': XL_CHART_TYPE.BAR_STACKED,
    '100% STACKED BAR': XL_CHART_TYPE.BAR_STACKED,
    'LINE': XL_CHART_TYPE.LINE,
    'TABLE': None,
    'STAT' : None,
    'PICTURE': None
}

chartstyles = {
    'blueberry': 3,
    'strawberry': 4,
    'mint': 5,
    'mandarin': 6,
    'sultana': 7,
    'grape': 8
}

brand_colors = {
    'default': MSO_THEME_COLOR.TEXT_1,
    'blueberry': MSO_THEME_COLOR.ACCENT_1,
    'strawberry': MSO_THEME_COLOR.ACCENT_2,
    'mint': MSO_THEME_COLOR.ACCENT_3,
    'mandarin': MSO_THEME_COLOR.ACCENT_4,
    'sultana': MSO_THEME_COLOR.ACCENT_5,
    'grape': MSO_THEME_COLOR.ACCENT_6
}

overall_list = ['OVERALL', 'SUM', 'TOTAL']  # If cat or series in this list, change color.

other_list = ['OTHER', 'SAME', 'PREFER NOT TO SAY', 'NEVER', 'OTHER: PLEASE SPECIFY']  # If cat or series in this list, change color.

highlight_list = []

# Dummy data to insert when major data errors arise.
data_error = {
    'categories': ['data mismatch', 'error', 'data mismatch', 'error'],
    'error': [0, 0, 0, 0]
}


# Separate from logging, places error notes in the deck
errordict = {
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

# Countries for country report
countrylist = [
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
    'United Arab Emirates',
    'United Kingdom',
    'United States'
]

# Exceptions for cell values
exceptionlist = ['-', '`', '*', '[', ']', 'u', 'N/A', 'NA']  # DELETE AFTER REFACTOR !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

exception_list = ['-', '`', '*', '**', '[', ']', 'u', 'N/A', 'NA']
exception_letter_string = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

# List of words to remove from titles
stopwords = {'PLEASE', 'FOODSERVICE', 'THE', 'WHAT', 'WHO', 'HOW', 'WHICH', 'WHEN', 'IS', 'ARE', 'CHECK', 'CHOOSE',
             'SELECT', 'OPTION', 'ONLY', 'MAIN', 'ONE', 'YOU', 'YOUR', 'MANY', 'DO', 'DOES', 'OF…', 'IF', 'THE',
             'ON', 'THAT', 'ALL', 'PART', 'FOLLOWING', 'BEST', 'DESCRIBES', 'APPLY', 'MARKET', 'OPERATION',
             'REPRESENT', 'THINKING', 'BACK', 'GENERALLY', 'SPEAKING', 'RESTAURANT', 'SERVES'
}

legend_locations = {
    'bottom': XL_LEGEND_POSITION.BOTTOM,
    'right': XL_LEGEND_POSITION.RIGHT
}

def monthlst(number_of_months, last_month, year=None, seperator='\n'):
    # Creates a list of months equal to length of provided number of months, ending with the last month
    # Adds years as needed
    month = last_month + 1
    month_deq = deque(month_abbr[1:])
    month_deq.rotate(1 - month)
    month_lst_full = list(month_deq)

    if year != None:
        year -= 1
        for m_idx, m in enumerate(month_lst_full):
            if 'JAN' in str(m).upper():
                year += 1
            month_lst_full[m_idx] = str(m) + seperator + str(year)

    month_lst = month_lst_full[(number_of_months * -1):]
    return month_lst


def titlecleaner(title):
    # splits string into list of words and removes punctuation
    txt = title
    for punc in ['?', '.', ',']:
        txt = txt.replace(punc,'')
    titlewords = txt.upper().split()
    print(titlewords)
    resultwords = [word for word in titlewords if word not in stopwords]
    try:
        if resultwords[-1] in ['IN', 'OF']:
            resultwords = resultwords[:-1]
        if resultwords[0] in ['IN', 'OF']:
            resultwords = resultwords[1:]
    except IndexError:
        pass

    # If title is empty, returns original (e.g., "Which One" to "")
    if len(resultwords) == 0:
        newtitle = title.upper()
    else:
        # Joins list back to string
        newtitle = ' '.join(resultwords)
    print(newtitle)
    return newtitle


# Returns new dictionary values each time.
def assign_chart_config(
        intendedchart='BAR', topbox=False, heatmap=False, highlight=False, indexes=None, preferredcolor='blueberry',
        banding='rows', title=None, legendloc='DEFAULT', cataxis=True, maxvals=[], growth=False, pct_dec_places=0,
        forceint=False, percentcheck=True, labeltxt=None, hasdata=False, tabcolor=None, state='visible', dq=[],
        note=[], datalabels=True):

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
        '*FORCE INT': forceint,
        'number format': '0.0%',
        'dec places': pct_dec_places,
        # Chart Styling
        '*TOP BOX': topbox,
        '*HEAT MAP': heatmap,
        '*HIGHLIGHT': highlight,
        'intended chart': intendedchart,
        'preferred color': preferredcolor,
        'banding': banding,
        'cat axis': cataxis,
        'chart title': title,
        'legend loc': legendloc,
        'max values': maxvals,
        'data labels': datalabels,
        # Page Styling
        '*INDEXES': indexes,
        '*GROWTH': growth,
        # Formatting checks and preflights
        'percent check': percentcheck,
        'error list': [],
        'notes': [],
        'chart chosen': False,
        'directional check': False,
        # Text Additions
        'title question': [],
        'bases': [],
        'data question': dq,
        'copy': [],
        'label text': labeltxt,
        'note': note,
        # Data reading/cleanup
        'has data': hasdata,
        'tab color': tabcolor,
        'sheet state': state
    }
    return chart_config

def assign_page_config(chart_count=0, table_count=0, tag=None):
    page_config = {
        'number of charts': chart_count,
        'number of tables': table_count,
        'has stat': False,
        'page title': None,
        'section tag': tag,
        'page copy': [],
        'footer': [],
        'page img': [],
        'function': None,
        'full page chart': False,
        'callouts': {}
    }
    return page_config


# Update Status

ui_errors = {
    'UnboundLocalError': 'Check Report Type',
    'PermissionError': 'File open, cannot save',
    'IndexError': 'Data Selection Error',
    'ValueError': 'Check Report Type'
}

tabname = []

def statusupdate(app, msg, fieldno):
    app.setStatusbar(msg, field = fieldno)
    app.setStatusbarBg("light gray", field = fieldno)
    app.topLevel.update()

def errorupdate(app, msg):
    for val in range(0, 4):
        app.setStatusbarBg("yellow", field=val)
        app.setStatusbar('', field=val)
    app.setStatusbar(msg, field=0)
    if msg in ui_errors:
        if msg == 'IndexError':
            ui_msg = ui_errors[msg]
            app.setStatusbar(str(tabname[-1]), field=2)
        else:
            ui_msg = ui_errors[msg]
        app.setStatusbar(ui_msg, field=1)
    else:
        app.setStatusbar('Contact Administrator', field=1)

    app.topLevel.update()



# Report-Specific Variables

idx_roundup_columns = {
    'lto' : [
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

scorecard_layouts = [
    'IGNITE_TitlePage',
    'TT—Subsection Intro, Main Ideas',
    'TT—Full Text',
    'TT_Primary_Table',
    'TT_Primary_Chart',
    'TT_6_Chart_&_Text',
    'TT_3_Chart_Dashboard_Flipped',
    'TT_Primary_Table_&_Text',
    'TT_End Wrapper_w_Photos'
]

dtv_layouts = [
    'TT_3_chart_2_table',
    'TT_Two_Chart_Equal',
    'TT_Primary_Table_&_Text'
]

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
                    'Household Income', 'Household Income', 'Household Income', 'Household Income', 'Household Income', 'Household Income',
                    'Ethnicity', 'Ethnicity', 'Ethnicity', 'Ethnicity'
                    ]

mock_chart = {
    'categories': ['Item 1', 'Item 2', 'Item 3', 'Item 4'],
    'Brand 1': [83, 167, 136, 80],
    'Brand 2': [83, 167, 136, 80],
    'Brand 3': [83, 167, 136, 80],
    'Brand 4': [83, 167, 136, 80],
}

# All employees than need to be imported into reports

employees = {
    'lh' : [
    'Lauren Hallow',
    'Senior Manager, Research and Insights',
    'lhallow@technomic.com'
    ],
    'jc' : [
    'Jenna Carroll',
    'Research Analyst',
    'jrcarroll@technomic.com'
    ]
}