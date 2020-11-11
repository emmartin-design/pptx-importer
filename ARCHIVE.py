# RETIRED. Chart.chart_style assigns colors now
# Pulls values from dictionary in Variables.py
def seriescolor_formatting2(chart, ic, catlen, slen, color, other=None, overall=None):
    if ic == 'PIE':
        for sc, series in enumerate(chart.series):
            for i, point in enumerate(series.points):
                bright_control = brightness_level(int(catlen), i)
                point = chart.series[sc].points[i]
                fill = point.format.fill
                fill.solid()
                fill.fore_color.theme_color = color  # Assigns the correct color
                fill.fore_color.brightness = bright_control
                if int(catlen) > 2:  # Overwrites color with slate or gray if certain terms used
                    if i == overall:
                        fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_1
                    elif i == other:
                        fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
                        fill.fore_color.brightness = 0.15
    else:
        for i, series in enumerate(chart.series):
            bright_control = brightness_level(int(slen), i)
            fill = series.format.fill
            fill.solid()
            fill.fore_color.theme_color = color  # Assigns the correct color
            fill.fore_color.brightness = bright_control
            if int(slen) > 2:  # Overwrites color with slate or gray if certain terms used
                if i == overall:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_1
                elif i == other:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
                    fill.fore_color.brightness = 0.15


# Cleaned up to look at number of charts regardless of single chart or no.
def layout_chooser(noc, templatedata, slidefunction = None, table = False):  # Doesn't work currently. Tweak
    chosen = None
    layoutholder = None
    for layout in templatedata:
        l = templatedata[layout]
        if slidefunction != None:  # Primarily used for non-general reports
            if slidefunction == 'cover':
                if 'TITLE 0' in l:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'intro':
                if l['bodycount'] == 6:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'text':
                if l['bodycount'] == 2 and l['titlecount'] == l['chartcount'] == l['tablecount'] == 0:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'scorecard':
                if l['bodycount'] == 2 and l['chartcount'] == 6 and l['preferred'] == True:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'table and text':
                if l['bodycount'] == 2 and l['chartcount'] == 0 and l['tablecount'] == 1:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'table':
                if l['bodycount'] == 1 and l['chartcount'] == 0 and l['tablecount'] == 1:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'endwrapper':
                if l['bodycount'] == 5 and l['picturecount'] == 3:
                    chosen, layoutholder = l, layout
            else:
                if noc == 1:
                    if table == True:
                        if l['tablecount'] == 1 and l['bodycount'] == 2:
                            chosen, layoutholder = l, layout
                    else:
                        if l['chartcount'] == 1:
                            chosen, layoutholder = l, layout
                else:  # Multiple chart layout chosen based on number of charts.
                    if table == True:
                        if l['tablecount'] > 0 and l['chartcount'] > 0:
                            if l['tablecount'] == noc and l['chartcount'] == noc:
                                chosen, layoutholder = l, layout
                        elif l['tablecount'] == noc:
                            chosen, layoutholder = l, layout
                    else:
                        if l['chartcount'] == noc:
                            chosen, layoutholder = l, layout


    return layoutholder, chosen



def add_table_line_graph(chart):
    chart_xml = chart._plotArea._element
    print(chart_xml)
    fld_xml = (
            '<c:dTable>\n'
            '    <c:showHorzBorder val="1"/>\n'
            '    <c:showVertBorder val="1"/>\n'
            '    <c:showOutline val="1"/>\n'
            '    <c:showKeys val="1"/>\n'
            '</c:dTable>' % nsdecls("a")
    )
    fld = parse_xml(fld_xml)
    chart_xml.append(fld)