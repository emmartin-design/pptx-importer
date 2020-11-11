from openpyxl import load_workbook
from collections import OrderedDict
import pandas as pd
import logging
from datetime import datetime

#Local Module
import variables as v

# defining variables
chartypelist = v.charttypelist
data_error = v.data_error
errordict = v.errordict
exceptionlist = v.exceptionlist

# pandas can open and return all worksheets as a dictionary.
# Use pandas to return data once the interface is built out.
def readsheet(file, wksht):
    df = pd.read_excel(file, sheet_name=wksht)
    return df

def longestval(lst):
    longest_val = lst[0]
    shortest_val = lst[0]
    for val in lst:
        if len(val) > len(longest_val):
            longest_val = val
        elif len(val) < len(shortest_val):
            shortest_val = val
    return longest_val, shortest_val


def strcleanup(string, upper=False):
    cleanstr = str(string).replace(':', '', 10)
    if upper:
        cleanstr = cleanstr.upper().strip()
    return cleanstr


def sheet_to_df(file):  # Used for position-based report types
    exceldata = pd.read_excel(file, index_col=0, sheet_name=0)
    df = pd.DataFrame(exceldata)
    textdf = pd.read_excel(file, sheet_name=1)
    textdata = textdf.iloc[:, 0].to_list()
    textdata.insert(0, textdf.columns[0])
    loggingmsg = 'Converted '  + file + ' to dataframe'
    logging.info(loggingmsg)
    return df, textdata


# Removes any blank entries from the data dictionary
def scrubber(dataog, app=None):
    data = dataog.copy()
    for slidecount, slide in enumerate(dataog):
        if app != None:
            msg_for_ui = 'Scrubbing Data — ' + str(round(((slidecount / len(dataog)) * 100), 0)) + '%'
            v.statusupdate(app, msg_for_ui, 1)
        page_config = dataog[slide]['page config']
        page_config['number of tables'] = 0
        page_config['number of charts'] = 0
        for card in dataog[slide]:
            if card != 'page config':
                config = dataog[slide][card]['config']
                if config['intended chart'] == 'TABLE':
                    page_config['number of tables'] += 1
                elif config['has data']:
                    page_config['number of charts'] += 1

        datavizcount = page_config['number of charts'] + page_config['number of tables']

        if page_config['page title'] != None:
            copycount = 1 + len(page_config['page copy'])
        else:
            copycount = len(page_config['page copy'])

        if datavizcount + copycount == 0:
            del data[slide]
            infolog('Scrubbed', data)
    return data

def infolog(msg, val):
    logging.info(str(msg)+ ": " + str(val))

def readbook(app, file, pptxname, country = None):
    wb = load_workbook(filename=file, data_only=True)
    slide_data = OrderedDict()

    combinecount = 1  # Indicates how many charts per slide
    most_recent_tabcolor = None

    slidecount = 0

    for sheetcount, wksht in enumerate(wb.worksheets):
        msg_for_ui = 'Reading Excel — ' + str(round(((sheetcount / len(wb.worksheets)) * 100), 0)) + '%'
        v.statusupdate(app, msg_for_ui, 1)
        infolog(wksht, wksht.sheet_state)
        if wksht.sheet_state == "visible":
            tabcolor = wksht.sheet_properties.tabColor
        if sheetcount not in slide_data:
            slide_data[sheetcount] = {}  # sets up dictionary for future use
            slide_data[sheetcount]['page config'] = v.assign_page_config(tag=pptxname)

        # Creates new config dict for each tab
        config = v.assign_chart_config(state=wksht.sheet_state, tabcolor=wksht.sheet_properties.tabColor)

        if wksht.sheet_state == "visible":
            tabcolor = wksht.sheet_properties.tabColor  # Used for combining charts

            config['notes'].append(wksht)
            framedata = OrderedDict()  # Used to create dataframe
            framedata['categories'] = []
            seriesnamelist = []
            colchecklst = []

            for col in wksht.iter_cols():  # Colcheckno is used to see if bases are vertical
                colcheckno = 0  # Resets to 0 each column. If higher than 0, indicates vertical positioning.

                serieslist = []  # blanks out series for every new column
                for cell in col:
                    cellval = wksht[cell.coordinate].value
                    cellcolor = cell.fill.start_color.index

                    # update config
                    cellval_c = strcleanup(cellval, True)  # Used for commands and chart types, not copy.
                    if cellval_c in config.keys():
                        config[cellval_c] = True
                    elif cell.font.underline == 'single':  # detects slide copy command
                        config['copy'].append(cellval)
                    elif cellval_c in chartypelist:  # detects chart type command
                        config['intended chart'], config['chart chosen'] = cellval_c, True

                    # color based selections
                    elif cellcolor == 4:  # detects data question for footer/chart title
                        config['title question'].append(cellval)  # Update split to question and title
                        infolog('DQ', cellval)

                    elif cellcolor == 5:  # detects bases for footer
                        colchecklst.append(colcheckno)  # Used to determine base labels (rows v. cols)
                        try:
                            config['bases'].append(('{:,}'.format(cellval)))  # Adds commas to base numbers
                        except ValueError:
                            config['bases'].append(cellval)
                        infolog('Base', cellval)
                        colcheckno += 1

                    elif cellcolor == 7:  # detects categories
                        if cellval is None:
                            config['error list'].append('106')
                            logging.warning('106 ERROR: blank category cell collected')
                        else:
                            cat = strcleanup(cellval)
                            framedata['categories'].append(cat)

                    elif cellcolor == 8:  # Collects series name (TO UPDATE: merge with series list with POP?)
                        if cellval is not None:
                            series_name = strcleanup(cellval)
                        else:
                            series_name = "ERROR_PLACEHOLDER"
                        seriesnamelist.append(series_name)

                    elif cellcolor == 9:  # detects data
                        config['has data'] = True
                        if config['*FORCE FLOAT'] or config['*FORCE CURRENCY']:
                            config['percent check'] = False
                        if cellval is None:  # Separate check for error reporting
                            serieslist.append(0)
                            config['error list'].append('106')
                            logging.warning('106 ERROR')

                        if config['*FORCE PERCENT'] is True:
                            if cellval in exceptionlist:
                                percentvalue = 0
                            else:
                                percentvalue = float(cellval) / 100
                            serieslist.append(percentvalue)
                        elif type(cellval) != float:
                            if type(cellval) is int and cellval > 1:
                                serieslist.append(float(cellval))
                            if cellval == 1 and config['percent check'] is True:
                                serieslist.append(1.0)
                            elif cellval == 0 or cellval_c in exceptionlist:
                                serieslist.append(0.0)
                            elif '=' in str(cellval):
                                config['error list'].append('107')  # TO UPDATE: Read formulas
                                logging.warning('107 ERROR')
                                serieslist.append(0.0)

                        else:
                            serieslist.append(cellval)

                if len(serieslist) > 0:  # Checks if series were collected
                    if len(framedata['categories']) != len(serieslist):
                        config['error list'].append('100')
                    framedata[series_name] = serieslist  # Adds series to data



            # Cleanup series names for vertical series (happens after initial collection)
            if len(seriesnamelist) > (len(framedata) - 1):
                newdata = OrderedDict()
                newdata['categories'] = framedata['categories'].copy()

                # The following appends non-repeating series names.
                # The set command doesn't maintain order, causing major issues.
                newseriesnames = []
                for name in seriesnamelist:
                    if name not in newseriesnames:
                        newseriesnames.append(name)
                lststep = len(newseriesnames)  # determines the step number for the data reorg

                # Find series name in original data for parsing
                if 1 < len(framedata) < 3:
                    for key in framedata:
                        if key != 'categories':
                            keyholder = key
                else:
                    logging.error('Major data issue on ' + str(wksht) + '. Check series selections.')

                for lstidx, series in enumerate(newseriesnames):
                    newdata[series] = []
                    for idx in range(lstidx, len(framedata[keyholder]), lststep):
                        newdata[series].append(framedata[keyholder][idx])

                framedata = newdata.copy()  # Replaces erroroneous data with reconfigured data
                if '100' in config['error list']:  # Removes previously applied error for vertical series charts
                    config['error list'].remove('100')

            # Creates a dataframe from the collected data. The dictionary is dropped from here on in.
            try:
                if country is not None:  # If this is a country report, it will filter out countries
                    newbases = []
                    df_og = pd.DataFrame.from_dict(framedata)
                    basedict = {'categories': 'bases'}
                    for idx, col in enumerate(df_og.columns, start=-1): # Create dictionary of bases
                        if col != 'categories':
                            basedict[col] = config['bases'][idx]
                    df_w_bases = df_og.append(basedict, ignore_index=True)
                    new_df = df_w_bases[['categories', 'Global Average', country]].copy()
                    baseholder = new_df.iloc[-1:].values.tolist()
                    for base in baseholder:
                        for b in base:
                            if b != 'bases':
                                newbases.append(b)
                    config['bases'] = newbases
                    new_df.drop(new_df.tail(1).index,inplace=True)
                    df = new_df
                    df.set_index('categories', inplace=True)
                else:
                    df = pd.DataFrame.from_dict(framedata)
                    df.set_index('categories', inplace=True)

                config['max values'] = df.max(axis=1).values.tolist()

            except:  # This drops in dummy data if the data doesn't fit into the dataframe.
                df = pd.DataFrame.from_dict(data_error)
                config['error list'].append('100')
                logging.warning('100 ERROR')
                df.set_index('categories', inplace=True)


            # Base Collection—Allows multiple bases
            data_base_list = []
            basecount = len(config['bases'])
            config['directional check'] = basecount > 0 and '35' in config['bases']
            colcheck = 1 in colchecklst  # True if vertical base selection
            basestr = ''  # Used for complex country report base formatting
            for nameidx, base in enumerate(config['bases']):
                if basecount == 1:
                    data_base_list.append('Base: ' + str(base))
                else:
                    if country is None:
                        if colcheck:  # Means bases refers to categories, not series
                            data_base_list.append(str(framedata['categories'][nameidx]) + ' Base: ' + str(base))
                        else:
                            data_base_list.append(str(df.columns[nameidx]) + ' Base: ' + str(base))

                    else:  # TO UPDATE: This will break is GA comes after country in data.
                        if df.columns[nameidx] == 'Global Average':
                            basestr = ('Base: ' + str(base) + '(global average)')
                        elif df.columns[nameidx] == country:
                            basestr = basestr + (' and ' + str(base) + '(' + str(df.columns[nameidx]) + ')')
                            data_base_list.append(basestr)

            config['bases'] = data_base_list

            if len(df.index) > 10 and config['intended chart'] == 'PIE':
                config['intended chart'] = 'BAR'

            #  Transposes Data
            if config['*TRANSPOSE'] is True:
                df_t = df.transpose()
                df = df_t  # replaces data with transposed version
                config['notes'].append("Chart Transposed")
                logging.info("Chart Transposed")

            # If chart is Top Box, create sum column
            if config['*TOP BOX'] == True:
                if config['intended chart'] not in ['STACKED BAR', 'STACKED COLUMN']:
                    config['intended chart'] = 'STACKED BAR'
                topboxname = 'Top ' + str(len(df.columns)) + ' Box'
                df[topboxname] = df.sum(axis=1)

            #  Sort data with Pandas. Does not work if category/series mismatch error thrown earlier
            if config['*SORT'] is True and '100' not in config['error list']:
                if len(framedata) > 0:
                    colname = df.columns[0]
                    s_note = "Chart sorted by " + str(colname)
                    config['notes'].append(s_note)
                    if config['*TOP BOX'] == True:
                        df_s = df.sort_values(by=[topboxname, colname], ascending=[False, False])
                    else:
                        df_s = df.sort_values(by=[colname], ascending=[False])
                    df = df_s

            #  Combine Check looks at tab color to see if new tab combined with previous on slide
            if combinecount > 6:
                logging.warning('Too many charts for combining. Cannot combine more than 6.')
                combinecount = 1
            if tabcolor is not None:
                if tabcolor == most_recent_tabcolor:  # compares tab to most recent tab
                    combinecount += 1
                    slidecount -= 1  # Keeps slide count accurate for error reporting
                else:
                    combinecount = 1
                most_recent_tabcolor = tabcolor
            else:
                combinecount = 1

            slide_data[slidecount][combinecount] = {'config': config, 'frame': df}
            slidecount += 1

    # Scrub Blanks
    msg_for_ui = 'Scrubbing Data'
    v.statusupdate(app, msg_for_ui, 1)
    slide_data_sd = scrubber(slide_data, app)
    msg_for_ui = 'Data Collection Complete'
    v.statusupdate(app, msg_for_ui, 1)

    return slide_data_sd


def scorecard(file, valrpt=False):
    exceldata = pd.read_excel(file, sheet_name=0)
    explainertext = pd.read_excel(file, sheet_name=1)
    df = pd.DataFrame(exceldata)
    slide_data = OrderedDict()
    splitdems = False

    # Set up data
    if valrpt == True:
        rpt_type = 'value'
        forepages = 4
        aftpages = 2
        rpt_title = 'Value Concept Screener Summary'
        scorecard_chartcount = 6
    else:
        rpt_type = 'lto'
        forepages = 5
        aftpages = 8
        rpt_title ='Menu Concept Screener Summary'
        scorecard_chartcount = 4
        explainertext2 = pd.read_excel(file, sheet_name=2)

    df.columns = v.scorecard_dfcols[rpt_type]
    df.set_index('Concept', inplace=True)

    if len(df.index) > 10:  # For the demographics table
        aftpages += 1
        splitdems = True
        if len(df.index) > 20:
            aftpages += 1
    totalpages = forepages + aftpages + len(df.index)

    # Set up pages
    for page in range(0, (totalpages)):
        slide_data[page] = {}  # Creates Page
        slide_data[page]['page config'] = v.assign_page_config()


    # Collect the data
    for page in range(0, (totalpages)):
        page_config = slide_data[page]['page config']
        if page == 0:  # Works for both LTO and Val LTO reports
            page_config['page title'] = rpt_title
            page_config['function'] = 'cover'
        else:
            page_config['section tag'] = rpt_title

            # Add big idea, small idea page
            if (page == 1 and rpt_type == 'value') or (page == 3 and rpt_type == 'lto'):
                page_config['page copy'] = ['Big Idea', 'Small Idea']
                page_config['function'] = 'intro'

            # Adds Descriptor slide #1
            elif (page == 2 and rpt_type == 'value') or (page == 1 and rpt_type == 'lto'):
                page_config['function'] = 'text'
                introtitle = explainertext.columns[0]
                text = explainertext[introtitle].values.tolist()
                text.insert(0, introtitle)
                page_config['page copy'] = text

            # If LTO, adds descriptor slide 2
            elif page == 2 and rpt_type == 'lto':  # Only works for non-value scorecards
                page_config['function'] = 'text'
                introtitle = explainertext2.columns[0]
                text = explainertext2[introtitle].values.tolist()
                text.insert(0, introtitle)
                page_config['page copy'] = text

            # Adds benchmarking table to both report types, different data maps
            elif page == (forepages - 1):  # idx roundup table page
                page_config['function'] = 'table'
                page_config['number of tables'] += 1
                page_config['full page chart'] == True
                slide_data[page][1] = {}
                if rpt_type == 'lto':
                    index_colnames = ['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability']
                    indexes_df = df[index_colnames].copy()
                    index_ranges = [[96, 122], [100, 127], [99, 116], [98, 117]]
                    for col in index_colnames:
                        indexes_df[col] = (indexes_df[col] * 100).astype(int)
                    slide_data[page][1]['frame'] = indexes_df
                elif rpt_type == 'value':
                    index_colnames = ['Purchase Intent Percentile', 'Value Percentile', 'Draw Percentile',
                                      'Craveability Percentile', 'Quality Percentile']
                    slide_data[page][1]['frame'] = df[index_colnames].copy()
                    index_ranges = None
                idx_roundup_config = v.assign_chart_config(indexes=index_ranges, intendedchart='TABLE', hasdata=True,
                                                    banding='cols')
                idx_roundup_config['*FORCE INT'] = True
                idx_roundup_config['percent check'] = False
                slide_data[page][1]['config'] = idx_roundup_config

            # Creates Scorecards
            elif (totalpages - aftpages) > page >= forepages:
                idx = (page - forepages)
                concept = df.index[idx]
                page_config['page copy'] = [('/tag' + str(df.iloc[idx,0]).upper()),
                                         ('/h1' + str(df.index[idx])), df.iloc[idx,1]]

                for count in range(1, (scorecard_chartcount + 1)):
                    slide_data[page][count] = {}

                if rpt_type == 'value':
                    # Stack Bar Charts
                    stacked_charts = {
                        'Value': ['Val 2nd Box', 'Val Top Box', 'Val Top 2 Box'],
                        'Purchase Intent': ['PI 2nd Box', 'PI Top Box', 'PI Top 2 Box'],
                    }
                    for stacked_idx, chart in enumerate(stacked_charts, start=1):
                        stacked_df = df.loc[[concept], stacked_charts[chart]].copy()
                        stacked_df.columns = ['2nd Box', 'Top Box', 'Top 2 Box']
                        stacked_df.index = [chart]
                        if stacked_idx == 1:
                            pref_color = 'blueberry'
                        else:
                            pref_color = 'mint'
                        slide_data[page][stacked_idx]['frame'] = stacked_df
                        slide_data[page][stacked_idx]['config'] = v.assign_chart_config(intendedchart='STACKED COLUMN', topbox=True,
                                                                       preferredcolor=pref_color, hasdata=True,
                                                                       legendloc='bottom', note=v.scorecard_footercopy[rpt_type])

                    # Draw, Craveability, Quality stacked bar
                    ddf = df.loc[[concept], ['Draw 2nd Box', 'Draw Top Box', 'Draw Top 2 Box']].copy()
                    cdf = df.loc[[concept], ['CRV 2nd Box', 'CRV Top Box', 'CRV Top 2 Box']].copy()
                    qdf = df.loc[[concept], ['Quality 2nd Box', 'Quality Top Box', 'Quality Top 2 Box']].copy()
                    for f in [ddf, cdf, qdf]:
                        f.columns = ['2nd Box', 'Top Box', 'Top 2 Box']
                    dcqdf = pd.concat([ddf, cdf, qdf])
                    dcqdf.index =  ['Draw', 'Craveability', 'Quality']

                    slide_data[page][3]['frame'] = dcqdf
                    slide_data[page][3]['config'] = v.assign_chart_config(intendedchart='STACKED COLUMN', topbox=True,
                                                                   preferredcolor='mint', hasdata=True,
                                                                   legendloc='bottom')

                    # Pie Charts
                    pie_charts = {
                        'Good or very good value reasoning': ['Portion', 'Taste', 'Quality', 'Fits Budget'],
                        'Order behavior': ['Only the offer', 'Additional foods and/or beverages', 'Unsure'],
                        'Repeat trial': ['Once', 'Some visits', 'Most visits', 'Every visit']
                    }

                    for pie_idx, pie in enumerate(pie_charts,start=4):
                        chart_title = pie
                        pie_df = df.loc[[concept], pie_charts[pie]].copy()
                        if pie_idx == 4:
                            pie_df['Other'] = 1 - pie_df.sum(axis=1)
                            pie_df.index = ['Base']
                            pref_color = 'blueberry'
                        else:
                            pref_color = 'mint'
                        pie_df_t = pie_df.transpose()
                        slide_data[page][pie_idx]['frame'] = pie_df_t
                        tempconfig = v.assign_chart_config(intendedchart='PIE', preferredcolor=pref_color,
                                                    title=chart_title, hasdata=True)
                        slide_data[page][pie_idx]['config'] = tempconfig

                    # Percenticle Indexes
                    callout_sz = [0.42, 0.3]
                    cols = ['Value Percentile', 'Purchase Intent Percentile', 'Draw Percentile',
                            'Craveability Percentile', 'Quality Percentile']
                    left_vals = [4.85, 8.02, 10.45, 11.33, 12.07]
                    page_config['callouts']['shape'] = [9.15, 0.3, 3.71, 1, 'white', 'PERCENTILES', 'l', False]
                    for left_val, col in zip(left_vals, cols):
                        page_config['callouts'][str(cols.index(col))] = callout_sz + [left_val, 1, None,
                                                                                   str(df.at[concept, col]),
                                                                                    'c', True]

                elif rpt_type == 'lto':
                    # PI, UNI, DRAW, CRV Stacked Bar
                    pidf = df.loc[[concept], ['PI 2nd Box', 'PI Top Box', 'PI Top 2 Box']].copy()
                    udf = df.loc[[concept], ['Uniqueness 2nd Box', 'Uniqueness Top Box', 'Uniqueness Top 2 Box']].copy()
                    ddf = df.loc[[concept], ['Draw 2nd Box', 'Draw Top Box', 'Draw Top 2 Box']].copy()
                    cdf = df.loc[[concept], ['CRV 2nd Box', 'CRV Top Box', 'CRV Top 2 Box']].copy()
                    for f in [pidf, udf, ddf, cdf]:
                        f.columns = ['2nd Box', 'Top Box', 'Top 2 Box']
                    pudcdf = pd.concat([pidf, udf, ddf, cdf])
                    pudcdf.index = ['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability']
                    slide_data[page][1]['frame'] = pudcdf
                    slide_data[page][1]['config'] = v.assign_chart_config(intendedchart='STACKED COLUMN', topbox=True,
                                                                   legendloc='bottom', hasdata=True)

                    # Willingness to pay stat
                    slide_data[page][2]['frame'] = [str(f"${df.at[concept, 'Median willingness to pay']:.2f}"),
                                                    'Median willingness to pay']
                    slide_data[page][2]['config'] = v.assign_chart_config(intendedchart='STAT', hasdata=True)
                    page_config['has stat'] = True

                    # Seasonality Pie
                    chart_title = 'Seasonality (Would purchase ______ during the year)'
                    sdf = df.loc[[concept], ['Once', 'Certain times', 'A few times']].copy()
                    sdf_t = sdf.transpose()
                    slide_data[page][3]['frame'] = sdf_t
                    slide_data[page][3]['config'] = v.assign_chart_config(intendedchart='PIE', title=chart_title, hasdata=True)

                    # Repeat Trial Pie
                    chart_title = 'Repeat trial'
                    rtdf = df.loc[[concept], ['Once', 'Some Visits', 'Most Visits', 'Every Visit']].copy()
                    rtdf_t = rtdf.transpose()
                    slide_data[page][4]['frame'] = rtdf_t
                    slide_data[page][4]['config'] = v.assign_chart_config(intendedchart='PIE', title=chart_title, hasdata=True)

                    # Percenticle Indexes
                    index_ranges = {1: [1.22, 0.96], 2: [1.27, 1.00], 3: [1.16, 0.99], 4: [1.17, 0.98]}
                    callout_sz = [0.5, 0.3]
                    cols = ['Purchase Intent', 'Uniqueness', 'Draw', 'Craveability']
                    left_vals = [4.7, 6.9, 9.06, 11.27]
                    page_config['callouts']['shape'] = [9.15, 0.3, 3.71, 1, 'white', 'INDEX', 'l', False, 'default']
                    for key_index, (left_val, col) in enumerate(zip(left_vals, cols), 1):
                        if df.at[concept, col] > index_ranges[key_index][0]:
                            text_color = 'mint'
                        elif df.at[concept, col] < index_ranges[key_index][1]:
                            text_color = 'mandarin'
                        else:
                            text_color = 'default'
                        page_config['callouts'][str(cols.index(col))] = callout_sz + [left_val, 1, None,
                                                                                   str(indexes_df.at[concept, col]),
                                                                                   'c', True, text_color]
            elif page == (totalpages - aftpages + 1):
                # demographics table(s)
                page_config['function'] = 'table and text'
                page_config['number of tables'] += 1
                demogsdf = df[['Male', 'Female', 'Gen Z', 'Millennials', 'Gen X', 'Baby Boomers',
                                                    '<$45K', '$45K-$99K', '$100K+', 'Black/African American',
                                                    'White', 'Hispanic/Latino']].copy()
                demogsdf_t = demogsdf.transpose()
                concepts = list(demogsdf_t.columns)

                for idx, concept in enumerate(concepts):
                    concepts[idx] = '/h2' +  str(idx + 1) + '. ' + concept

                demogsdf_t.columns = list(range(1, len(concepts) + 1))
                maxvals = demogsdf_t.max(axis=1).values.tolist()

                df_lst, concept_lst = [], []
                col_len = len(demogsdf_t.columns)
                if 21 > col_len > 10:
                    divide_col_len = int((col_len / 2))
                elif col_len > 20:
                    divide_col_len = int((col_len / 3))
                else:
                    divide_col_len = col_len
                for col in range(0, col_len, divide_col_len):
                    df_lst.append(demogsdf_t.iloc[:, col:col + divide_col_len])
                    concept_lst.append(concepts[col:col + divide_col_len])

                page_mod = 0  # Creates new page for split table if needed
                for frame_idx in df_lst:
                    if splitdems == True and page_mod > 0:
                        slide_data[(page + page_mod)]['page config'] = slide_data[page]['page config'].copy()
                    sd = slide_data[(page + page_mod)]
                    sd['page config']['page copy'] = concept_lst[page_mod]
                    sd[1] = {
                        'frame': frame_idx,
                        'config': v.assign_chart_config(maxvals=maxvals, intendedchart='TABLE', hasdata=True,
                                                     banding='cols', highlight=True)
                    }
                    page_mod += 1

            elif page == (totalpages - aftpages):
                page_config['number of charts'] = 1
                mock_df = pd.DataFrame.from_dict(v.mock_chart)
                mock_df.set_index('categories', inplace=True)
                sd = slide_data[page]
                sd[1] = {}
                sd[1]['frame'] = mock_df
                sd[1]['config'] = v.assign_chart_config(intendedchart='COLUMN', title='BRAND FIT BY ITEM', hasdata=True,
                                                 forceint=True, percentcheck=False)

            elif page == (totalpages - 1):  # Last page, -1 to compensate for zero-idxing
                page_config['function'] = 'endwrapper'
                page_config['page copy'] = ["So. What's Next?", "Need some more LTO guidance? Reach out to our experts."]
                sd = slide_data[page]

                for chart in [1, 2, 3, 4]:
                    sd[chart] = {}
                    if chart < 3:
                        ee = 'lh'
                    else:
                        ee = 'jc'
                    if (chart % 2) == 0:
                        sd[chart]['frame'] = v.employees[ee]
                        sd[chart]['config'] = v.assign_chart_config(intendedchart='STAT')
                    else:
                        sd[chart]['frame'] = 'templates/import_resources/headshots/' + ee + '.jpg'
                        sd[chart]['config'] = v.assign_chart_config(intendedchart='PICTURE')
    slide_data_sd = scrubber(slide_data)

    return slide_data_sd


def dtv_reader(file):
    exceldata = pd.read_excel(file, sheet_name=0)
    df = pd.DataFrame(exceldata)
    slide_data = OrderedDict()
    df.columns = v.scorecard_dfcols['dtv']
    df.set_index(['Account', 'Competitor No.'], inplace=True)

    title_months = v.monthlst(3, ((datetime.now().month) - 1))
    reportdate = str(title_months[-1]) + ' ' + str(datetime.now().year)

    page = 1
    for account, account_df in df.groupby(level=0):
        for layout in [1, 2, 3]:
            slide_data[page] = {}
            slide_data[page]['page config'] = v.assign_page_config()
            page_config = slide_data[page]['page config']
            page_config['section tag'] = str(account) + ' | Consumer Visit Tracker & Ignite Consumer | ' + reportdate
            if layout == 1:  # First page of account
                for chart in range(1, 6):
                    slide_data[page][chart] = {}
                page_config['number of charts'] = 3
                page_config['number of tables'] = 2
                page_config['function'] = 'dtv'

                # Sales Market Share Stacked Bar
                sms_df = account_df.loc[:, ['Competitor Brand', 'Sales Market Share']].copy()
                sms_df.set_index(['Competitor Brand'], inplace=True)
                sms_df_t = sms_df.transpose()

                # Set up values to replace data labels
                growthdict = {1:[], 2:[], 3:[], 4:[]}
                curr_df=account_df.loc[:, ['Sales Market Share']].copy()
                currtxt = curr_df['Sales Market Share'].values.tolist()
                smsg_df = account_df.loc[:, ['Chg Sales Market Share']].copy()
                growthtxt = smsg_df['Chg Sales Market Share'].values.tolist()
                for idx, item in enumerate(growthtxt):
                    currval = round((currtxt[idx] * 100), 1)
                    f = '(%s)' % round(growthtxt[idx], 1)
                    growthdict[idx] = [(str(currval) + '% ' + f)]

                # Creat chart title
                sharetitle = 'SALES MARKET SHARE\n'
                for month in title_months:
                    if month != title_months[-1]:
                        sharetitle += (str(month) + '-')
                sharetitle += (' ' + str(datetime.now().year) + ' VS. ' + str(datetime.now().year - 1))

                notetxt = '* Arrows indicate change from July YTD 2020 to August 2020 YTD'

                slide_data[page][1]['frame'] = sms_df_t
                slide_data[page][1]['config'] = v.assign_chart_config(intendedchart='100% STACKED BAR', legendloc='bottom',
                                                               title=sharetitle, cataxis=False, hasdata=True,
                                                               pct_dec_places=1, labeltxt=growthdict,
                                                               note=[notetxt])

                # YOY Tables
                table_titles = {
                    'sales': ['Competitor Brand', 'YOY Sales Qtr', 'YOY Sales Annual'],
                    'traffic': ['Competitor Brand', 'YOY Traffic Qtr', 'YOY Traffic Annual']
                }

                for t_idx, table in enumerate(table_titles, 2):
                    table_df = account_df.loc[:, table_titles[table]].copy()
                    table_df.set_index(['Competitor Brand'], inplace=True)
                    slide_data[page][t_idx]['frame'] = table_df
                    slide_data[page][t_idx]['config'] = v.assign_chart_config(intendedchart='TABLE', hasdata=True, growth=True,
                                                                   pct_dec_places=1)

                # Column Charts
                column_charts = {
                    'KEY PERFORMANCE INDICATORS*': {
                        'cols': ['Competitor Brand', 'Past Week Trial', 'Redeemed Coupon', 'Ordered LTO',
                                 'Lapsed User'],
                        'pp_cols': ['Competitor Brand', 'Chg Past Week Trial', 'Chg Redeemed Coupon', 'Chg Ordered LTO',
                                    'Chg Lapsed User'],
                        'color': 'mint'

                    },
                    'ORDERING METHOD*': {
                        'cols': ['Competitor Brand', 'Order In Store', 'Order Drive thru', 'Order Pickup/Delivery'],
                        'pp_cols': ['Competitor Brand', 'Chg Order In Store', 'Chg Order Drive thru',
                                    'Chg Order Pickup/Delivery'],
                        'color': 'grape'
                    }
                }

                for ch_idx, chart in enumerate(column_charts, 4):
                    ch_df = account_df.loc[:, column_charts[chart]['cols']].copy()
                    ch_pp_df = account_df.loc[:, column_charts[chart]['pp_cols']].copy()
                    for frame in [ch_df, ch_pp_df]:
                        frame.set_index(['Competitor Brand'], inplace=True)

                    ch_lst = ch_df.transpose().values.tolist()
                    ch_pp_lst = ch_pp_df.transpose().values.tolist()
                    ch_dict = {}
                    for s_idx, series in enumerate(ch_pp_lst):
                        ch_dict[s_idx] = []
                        for idx, item in enumerate(ch_pp_lst[s_idx]):
                            ch_val = int((ch_lst[s_idx][idx] * 100))
                            if ch_pp_lst[s_idx][idx] > 0:
                                growthval = u"\u2191"
                            elif ch_pp_lst[s_idx][idx] < 0:
                                growthval = u"\u2193"
                            else:
                                growthval = ''
                            ch_dict[s_idx].append((growthval + u"\u000A" + str(ch_val) + '%'))

                    slide_data[page][ch_idx]['frame'] = ch_df
                    slide_data[page][ch_idx]['config'] = v.assign_chart_config(intendedchart='COLUMN', legendloc='bottom',
                                                                   title=chart, hasdata=True, labeltxt=ch_dict,
                                                                   preferredcolor=column_charts[chart]['color'])

                # Chart Titles as callouts
                callout_sz = [2.71, .22]
                cols = ['YEAR-OVER-YEAR SALES', 'YEAR-OVER-YEAR TRAFFIC']
                left_vals = [3.62, 8.44]
                for left_val, col in zip(left_vals, cols):
                    page_config['callouts'][str(cols.index(col))] = callout_sz + [left_val, 1.19, None, col, 'l', True]

            elif layout == 2: # Line Graph Page
                line_graphs = {
                    'sales': {
                        'cols':['Competitor Brand', 'SM1','SM2','SM3','SM4','SM5','SM6','SM7'],
                        'title': 'SALES PERFORMANCE\nROLLING THREE-MONTH SYSTEMWIDE'
                    },
                    'traffic': {
                        'cols': ['Competitor Brand', 'TM1','TM2','TM3','TM4','TM5','TM6','TM7'],
                        'title': 'TRAFFIC PERFORMANCE\nROLLING THREE-MONTH SYSTEMWIDE'
                    }
                }

                for g_idx, graph in enumerate(line_graphs, 1):
                    slide_data[page][g_idx] = {}
                    g_df = account_df.loc[:, line_graphs[graph]['cols']].copy()
                    g_df.columns = ['Competitor Brand'] + v.monthlst(7, ((datetime.now().month) - 1), (datetime.now().year))
                    g_df.set_index(['Competitor Brand'], inplace=True)
                    g_df_t = g_df.transpose()
                    slide_data[page][g_idx]['frame'] = g_df_t
                    slide_data[page][g_idx]['config'] = v.assign_chart_config(intendedchart='LINE', legendloc=None,
                                                                        title=line_graphs[graph]['title'], hasdata=True,
                                                                       datalabels=False)

            elif layout == 3:  # Third page of account
                page_config['number of charts'] = 0
                page_config['number of tables'] = 1
                page_config['function'] = 'table and text'

                # Provides Entertainment Body Text (Segment level only)
                bodycopy = ['/b/h2Provides Video/TV Entertainment', 'Top Box', '']
                tv_df = df.loc[['Total'], ['Competitor Brand', 'Provides TV/Entertainment']].copy()
                tv_df.set_index(['Competitor Brand'], inplace=True)
                for idx, competitor in enumerate(tv_df.index):
                    val_copy = str(round((df.iloc[idx, 1] * 100), 1))
                    bodycopy.append(('/h1' + val_copy + '%'))
                    bodycopy.append(('/h2' + competitor))
                    bodycopy.append((''))
                page_config['page copy'] = bodycopy

                # Consumer Metrics Table
                slide_data[page][1] = {}
                cmt = { 'col_lst': ['Competitor Brand', 'Overall Ambience/Atmosphere', 'Video/TV Entertainment',
                                  'This was the right place for the occasion',
                                  'Appropriateness for the variety of occasions'],
                        'chg_col_lst': ['Competitor Brand', 'Chg Overall Ambience/Atmosphere',
                                      'Chg Video/TV Entertainment', 'Chg This was the right place for the occasion',
                                      'Chg Appropriateness for the variety of occasions']
                }


                cm_df = account_df.loc[:, cmt['col_lst']].copy()
                chg_cm_df = account_df.loc[:, cmt['chg_col_lst']].copy()
                for frame in [cm_df, chg_cm_df]:
                    frame.set_index(['Competitor Brand'], inplace=True)
                chg_cm_df_t = chg_cm_df.transpose()
                cm_df_t = cm_df.transpose()
                slide_data[page][1]['frame'] = cm_df_t
                slide_data[page][1]['config'] = v.assign_chart_config(intendedchart='TABLE', hasdata=True, indexes=chg_cm_df_t,
                                                               pct_dec_places=1)

            page += 1

    slide_data_sd = scrubber(slide_data)
    return slide_data_sd



