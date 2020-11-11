from openpyxl import load_workbook
import pandas as pd
import logging

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


def str_cleanup(string, upper=False):
    cleanstr = str(string)
    for mark in ["'", ":"]:
        cleanstr = cleanstr.replace(mark, "", 10)
    if upper:
        cleanstr = cleanstr.upper().strip()
    return cleanstr


def example_mover(text):
    open_par_idx = text.find('(')
    close_par_idx = text.find(')')
    new_text = text
    note = ''
    if open_par_idx > 0:
        new_text = text.replace(text[open_par_idx:(close_par_idx + 1)], '')
        note = new_text.strip() + '—' + text[(open_par_idx + 1):close_par_idx].strip()
    new_text = str_cleanup(new_text)
    return new_text.strip(), note


def sheet_to_df(file):  # Used for position-based report types
    exceldata = pd.read_excel(file, index_col=0, sheet_name=0)
    df = pd.DataFrame(exceldata)
    textdf = pd.read_excel(file, sheet_name=1)
    textdata = textdf.iloc[:, 0].to_list()
    textdata.insert(0, textdf.columns[0])
    loggingmsg = 'Converted '  + file + ' to dataframe'
    logging.info(loggingmsg)
    return df, textdata


def countryfy(df, country, bases):
    # Create new dataframe containing only the country for each report and the global average

    base_idx = df.columns.get_loc(country)
    new_bases = bases[base_idx]
    countryfied_df = df[['Global Average', country]].copy
    return(countryfied_df, new_bases)


def clean_up_data(data_dict, config):  # Checks data for gaps, formats based on config
    formatted_data_dict = {}
    series_names = []  # Used only for vertical series
    category_name_filter = ['Base', 'Median']

    # Checks if values should not be formatted as percentages
    for fig in ['*MEAN', '*FORCE CURRENCY', '*FORCE INT']:
        if config[fig] == True:
            config['percent check'] = False

    if not config['*MEAN']:
        category_name_filter.append('Mean')

    # Removes blank keys from dictionaries
    data_dict_no_blanks = {key:value for (key,value) in data_dict.items() if len(value) > 0}

    # Filter garbage out of data by key in data dict
    for series_idx, series in enumerate(data_dict_no_blanks):
        if series_idx == 0:  # First collected is always 'categories'
            data_dict_no_blanks[series].insert(0, 'categories')  # Used to even out length of columns
            for word in category_name_filter:
                try:
                    data_dict_no_blanks[series].remove(word)  # If "Base" is in categories, will remove
                except:
                    pass
                # This cleans up the category names
                for cat_idx, category in enumerate(data_dict_no_blanks[series]):
                    data_dict_no_blanks[series][cat_idx], notes_addition = example_mover(category)
                    if notes_addition != '':
                        config['notes'].append(notes_addition)
        else:
            garbage_collector = []  # Collects garbage values for later deletion

            # Detect vertical series and pull out set of names.
            if all(type(value) == str for value in data_dict_no_blanks[series]):
                [series_names.append(name) for name in data_dict_no_blanks[series] if name not in series_names]
                config['vertical series'] = True
            try:
                series_names.remove('Base')
            except:
                pass

            # Examine and collect garbage values
            for value_idx, value in enumerate(data_dict_no_blanks[series]):
                if type(value) == str:
                    # Filter out all uppercase strings or symbol strings
                    if value in v.exception_letter_string or value.isupper():
                        garbage_collector.append(value)
                    elif value in v.exception_list:
                        data_dict_no_blanks[series][value_idx] = 0.0
                else:
                    # Collect non-percent values
                    if config['percent check'] is True:
                        if value >= 1:
                            garbage_collector.append(value)
                    # If means, integers or currencies are needed, will filter out percents
                    else:
                        if value <= 1 or type(value) == int:
                            garbage_collector.append(value)

            # New series from values not in garbage collector. Deleting outright caused index errors
            data_dict_no_blanks[series] = [value for value in data_dict_no_blanks[series]
                                           if value not in garbage_collector]

        # Create new dictionary with correct key names
        if config['vertical series'] is False:
            formatted_data_dict[data_dict_no_blanks[series].pop(0)] = data_dict_no_blanks[series]

    # Split and reorder data for vertical series
    if config['vertical series'] is True:
        for series_idx, series in enumerate(data_dict_no_blanks):
            if series_idx == 0:
                formatted_data_dict['categories'] = data_dict_no_blanks[series]
            elif series_idx == len(data_dict_no_blanks) - 1:
                for name_idx, name in enumerate(series_names):
                    formatted_data_dict[name] = data_dict_no_blanks[series][name_idx::len(series_names)]

    return formatted_data_dict, config


def configure_pages(page_data, tag=None):
    for page in page_data:
        chart_count, table_count = 0, 0
        for chart in page_data[page]:
            if chart != 'config':
                if page_data[page][chart]['config']['intended chart'] == 'TABLE':
                    table_count += 1
                else:
                    chart_count += 1
        page_data[page]['page config'] = v.assign_page_config(chart_count, table_count, tag)
    return page_data


def infolog(msg, val):
    logging.info(str(msg)+ ": " + str(val))


def readbook(app, file, pptxname, country = None):
    wb = load_workbook(filename=file, data_only=True)
    page_data = {}
    most_recent_tab_color = None
    page_counter = 0

    sheet_data = {}
    for sheet_idx, sheet in enumerate(wb.worksheets):
        msg_for_ui = 'Reading Excel — ' + str(round(((sheet_idx / len(wb.worksheets)) * 100), 0)) + '%'
        v.statusupdate(app, msg_for_ui, 1)
        infolog(sheet, sheet.sheet_state)

        # Create default config data for chart and add it to the sheet data dictionary
        sheet_data[sheet_idx] = {}
        config = v.assign_chart_config(state=sheet.sheet_state, tabcolor=sheet.sheet_properties.tabColor)
        config['notes'].append(sheet)

        # Create dictionary to collect data. Can only be one 'categories' per sheet
        framedata = {}
        column_list = []  # Used for bases

        # Iterate over excel by column to collect data
        for col_idx, col in enumerate(sheet.iter_cols()):
            series_lst = []  # Resets the list for new series
            cell_strike_list = []  # Tracks data detection to help filter trash values

            # Check each cell for commands or color coding
            for cell_idx, cell in enumerate(col):
                cellval, cellcolor = sheet[cell.coordinate].value, cell.fill.start_color.index
                cellval_cleaned_up = str_cleanup(cellval, True)  # Removes extra spaces, capitalizes strings

                if cellval_cleaned_up in config.keys():  # Checks for commands
                    config[cellval_cleaned_up] = True

                elif cellval_cleaned_up in chartypelist:  # detects chart type
                    config['intended chart'], config['chart chosen'] = cellval_cleaned_up, True

                elif cellcolor == 4:  # detects data question for footer/chart title
                    config['title question'].append(cellval)  # Update split to question and title

                elif cellcolor == 5:  # detects bases for footer
                    column_list.append(col_idx)  # Tracks column indexes of bases to align them with correct axis
                    try:
                        config['bases'].append(('{:,}'.format(cellval)))  # Adds commas to base numbers when possible
                    except ValueError:
                        config['bases'].append(cellval)

                elif cellcolor == 7:  # detects all data
                    # First, does the cell have content?
                    if cellval is not None:
                        if cellval == 0:
                            series_lst.append(0.0)
                        elif cellval == 1:
                            cell_strike_list.append(0)
                        elif cellval_cleaned_up in v.exception_list:
                            if '*' in cellval_cleaned_up:
                                config['directional check'] = True
                            # If not a percent, is it far enough away from others to be counted as a null?
                            if (cell_idx - cell_strike_list[-1]) > 3:
                                series_lst.append(0)
                                cell_strike_list.append(0)
                        else:
                            series_lst.append(cellval)
                            cell_strike_list.append(cell_idx)

            framedata[col_idx] = series_lst  # Data is still pretty dirty at this point.

        # Cleans up collected data
        clean_data, config = clean_up_data(framedata, config)

        if len(clean_data) > 0:
            try:
                # Now that data is clean, create dataframe
                df = pd.DataFrame.from_dict(clean_data)
                df.set_index('categories', inplace=True)

                # Selects only specific country data for country reports
                if country is not None:
                    df, config['bases'] = countryfy(df.copy(), country, config['bases'].copy())
            except:
                # Data length mismatch causes errors
                error_data = {"ERROR":['CHECK', 'DATA', 'SHEET'], "COLLECTION":[0, 0, 0], "CHECK DATA":[0, 0, 0]}
                df = pd.DataFrame.from_dict(error_data)
                df.set_index('ERROR', inplace=True)
                config['error list'].append('100')
                config['notes'].append(str(clean_data))
        else:
            continue  # Stops iteration before recording empty data frames

        # Applies transposition
        if config['*TRANSPOSE']:
            df_t = df.transpose()
            df = df_t  # replaces data with transposed version
            config['notes'].append("Chart Transposed")

        # If chart is Top Box, create sum column
        if config['*TOP BOX']:
            if config['intended chart'] not in ['STACKED BAR', 'STACKED COLUMN']:
                config['intended chart'] = 'STACKED BAR'
            top_box_name = 'Top ' + str(len(df.columns)) + ' Box'
            df[top_box_name] = df.sum(axis=1)

        # Sort data by first column in DF after transposing, or if Top Box, sum first, then first column
        if config['*SORT']:
            config['notes'].append("Chart sorted by " + str(df.columns[0]))
            if config['*TOP BOX']:
                df_s = df.sort_values(by=[top_box_name, df.columns[0]], ascending=[False, False])
            else:
                df_s = df.sort_values(by=[df.columns[0]], ascending=[False])
            df = df_s

        # Used to highlight highest values in table row after transposing
        config['max values'] = df.max(axis=1).values.tolist()

        # Check Tabcolor against most recent. If different, create new page.
        if config['tab color'] is None or config['tab color'] != most_recent_tab_color:
            page_counter += 1
            page_data[page_counter] = {}

        most_recent_tab_color = config['tab color']

        page_data[page_counter][sheet_idx] = {
            'frame': df,
            'config': config
        }

    # Count data visualizations and assign page config
    formatted_page_data = configure_pages(page_data, tag=pptxname)

    return formatted_page_data





