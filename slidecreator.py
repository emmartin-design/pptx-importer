from pptx.util import Inches, Pt
from pptx import Presentation
from datetime import datetime, date, time
from appJar import gui
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import ChartData
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.slide import SlideLayout
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.xmlchemy import OxmlElement
import pandas as pd
import logging
import chartcreator as chc
import excelreader as er
import variables as v



def notesinsert(slide, notestext):
    try:
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        text_frame.text = str(notestext)
    except:
        pass


def preflightaddshape(slide, msg, lvl, boxplacement):
    left = top = Inches(boxplacement)
    width = height = Inches(2.5)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    fill = shape.fill
    fill.solid()
    if lvl == 'high':
        fill.fore_color.rgb = RGBColor(255, 0, 0)
    else:
        fill.fore_color.rgb = RGBColor(255, 255, 0)
    line = shape.line
    line.fill.background()
    shape.shadow.inherit = False
    shape.text = msg
    shape.text_frame.paragraphs[0].font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    shape.text_frame.paragraphs[0].font.color.size = Pt(10)


def preflight(slide, el):  # Inserts applicable error messages as boxes
    if len(el) > 0:
        for chart in el:
            if len(chart) == 0:
                logging.info("No Errors Found")
            else:
                boxplacement = 1.0
                for error in chart:
                    if error in v.errordict:
                        preflightaddshape(slide, v.errordict[error][0], v.errordict[error][1], boxplacement)
                        logging.warning(v.errordict[error][0])
                    boxplacement += 0.5


def add_gradient_legend(slide):
    width, height, top = Inches(1.35), Inches(0.33), Inches(0.47)
    legend_text = ['Low Index', 'High Index']
    left_loc = [10.13, 11.47]

    for text, loc in zip(legend_text, left_loc):
        left = Inches(loc)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        text_frame = shape.text_frame
        p = text_frame.paragraphs[0]
        p.text = text
        p.font.name = 'Arial'
        p.font.size = Pt(11)
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        shape.line.fill.background()
        shape.shadow.inherit = False
        fill = shape.fill
        fill.gradient()
        fill.gradient_angle = 0
        gradient_stops = fill.gradient_stops
        if 'Low' in text:
            gradient_stops[0].color.rgb = RGBColor(236, 124, 37)
            gradient_stops[0].position = 0.25
            gradient_stops[1].color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
            p.alignment = PP_ALIGN.LEFT
        else:
            gradient_stops[0].color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
            gradient_stops[1].color.rgb = RGBColor(163, 119, 182)
            gradient_stops[1].position = 0.25
            p.alignment = PP_ALIGN.RIGHT


def create_footer(slide, footercopy, directionalcheck=False):
    footer = None
    pageno = None
    for shape in slide.placeholders:
        if "SLIDE_NUMBER" in str(shape.placeholder_format.type):  # Master must have placeholder to function
            pageno = slide.placeholders[shape.placeholder_format.idx]
        elif "FOOTER" in str(shape.placeholder_format.type):  # Master must have placeholder to function
            footer = slide.placeholders[shape.placeholder_format.idx]

    footer_text_frame = footer.text_frame
    footer_pageno = pageno.text_frame.paragraphs[0]._p
    # edits XML directly—not ideal, but can't be updated until Python-pptx adds support.
    fld_xml = (
        '<a:fld %s id="{1F4E2DE4-8ADA-4D4E-9951-90A1D26586E7}" type="slidenum">\n'
        '  <a:rPr lang="en-US" smtClean="0"/>\n'
        '  <a:t>2</a:t>\n'
        '</a:fld>\n' % nsdecls("a")
    )
    fld = parse_xml(fld_xml)
    footer_pageno.append(fld)
    footer_a = 'Note: Due to small base, data is directional'
    paragraph_strs = footercopy
    if directionalcheck is True:
        paragraph_strs.append(footer_a)

    for para_str in paragraph_strs:
        p = footer_text_frame.add_paragraph()
        p.text = para_str.replace("'", "")


def footer_patch(self):
    for ph in self.placeholders:
        yield ph


SlideLayout.iter_cloneable_placeholders = footer_patch  # replaces module code with redefined latent placeholder code


def layout_chooser(chartcount, tablecount, has_stat, templatedata, slidefunction = None):  # Doesn't work currently. Tweak
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
            elif slidefunction == 'table and text':
                if l['bodycount'] == 2 and l['chartcount'] == 0 and l['tablecount'] == 1:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'table':
                if l['bodycount'] == 1 and l['chartcount'] == 0 and l['tablecount'] == 1:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'endwrapper':
                if l['bodycount'] == 5 and l['picturecount'] == 3:
                    chosen, layoutholder = l, layout
            elif slidefunction == 'dtv':
                if l['bodycount'] == 1 and l['chartcount'] == 3 and l['tablecount'] == 2:
                    chosen, layoutholder = l, layout
        else:
            if has_stat:  # Stats are considered charts, so chartcount is reduced by one to avoid too many placeholders
                if l['bodycount'] == 3 and l['tablecount'] == tablecount and l['chartcount'] == (chartcount - 1):
                    chosen, layoutholder = l, layout
            else:
                if l['tablecount'] == tablecount and l['chartcount'] == chartcount:
                    chosen, layoutholder = l, layout
                elif l['chartcount'] > 4 and l['chartcount'] == chartcount + 1:
                    chosen, layoutholder = l, layout


    return layoutholder, chosen


def assignchartdata(df, config, pct_dec_places = 1):
    if 'DataFrame' in str(type(df)):
        chart_data = ChartData()
        chart_data.categories = df.index
        percent_format = '0'
        if pct_dec_places > 0:
            percent_format += '.'
            for place in range((pct_dec_places)):
                percent_format += '0'
        percent_format += '%'
        config['number format'] = percent_format

        # Move to Data Cleanup
        if config['*FORCE FLOAT'] or config['*MEAN']:
            config['number format'] = '0.0'
        elif config['*FORCE INT']:
            config['number format'] = '0'
        elif config['*FORCE CURRENCY']:
            config['number format'] = '$0.00'

        for col in df.columns:
            chart_data.add_series(col, df[col], config['number format'])

    else:
        chart_data = df
    return chart_data


def scrub_formatting(text):
    scrubbedtext = str(text)
    for format in ['/b', '/i','/h1', '/h2', '/q', '/tag']:
        scrubbedtext = scrubbedtext.replace(format, '')
    return scrubbedtext


def insert_text(textlst, text_frame):
    for idx, para in enumerate(textlst):
        if idx < 1 and '/tag' not in para:
            p = text_frame.paragraphs[0]
            p.level = 0
        else:
            p = text_frame.add_paragraph()
            p.level = 1

        # Command-based formatting
        if '/b' in para:
            p.font.bold = True
        if '/tag' in para:
            p.font.color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
            p.font.color.brightness = -0.5
        if '/h1' in para:
            p.level = 0
        if '/h2' in para:
            p.level = 1
        p.text = scrub_formatting(para)


def callout_formatting(shape, run_text, alignment=PP_ALIGN.CENTER, indexing=False, text_color = 'default'):
    text_frame = shape.text_frame
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = text_frame.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = run_text
    font = p.font
    font.size = Pt(11)
    font.bold = text_color != 'default'
    font.color.theme_color = v.brand_colors[text_color]
    if text_color == 'mint':
        font.color.brightness = -0.25


def data_import(app, template, wbdata, templatedata, trusave, country = None, msg='Processing Data. '):
    logging.info('Starting Data Import')
    global prs
    prs = Presentation(template)
    # Select Layout and drop in charts
    for slideidx, page in enumerate(wbdata):
        try:
            pagemeta = wbdata[page]['page meta']
        except:
            pagemeta = wbdata[page]['page config']
        chartcount = pagemeta['number of charts']
        tablecount = pagemeta['number of tables']
        has_stat = pagemeta['has stat']
        function = pagemeta['function']
        footertext = []
        title_txt = pagemeta['page title']
        section_tag = pagemeta['section tag']

        logging.info('Slide no. ' + str(slideidx + 1) + '--There is/are ' + str(chartcount + tablecount) + ' Charts:')
        layout, shapeslist = layout_chooser(chartcount, tablecount, has_stat, templatedata, slidefunction=function)
        used_ph, used_data = [], []  # Used for multiple charts. Checks if placholders/data used already
        slide = prs.slides.add_slide(prs.slide_layouts[layout])
        slidenotes, preflightlst = [], []

        try:  # If page doesn't have a title, drops in section tags
            slide.shapes.title.text = title_txt
        except:
            for shape in shapeslist:
                if 'BODY' in shape:
                    if 12 > shapeslist[shape]['width'] > 9:
                        if shapeslist[shape]['height'] < 1:
                            phidx = shapeslist[shape]['index']
                            placeholder = slide.placeholders[phidx]
                            text_frame = placeholder.text_frame
                            p = text_frame.paragraphs[0]
                            p.text = section_tag

        for chart in wbdata[page]:
            if chart != 'page config':
                chart_config = wbdata[page][chart]['config']
                intendedchart = chart_config['intended chart']
                title_question = chart_config['title question']
                error_list = chart_config['error list']


                slidenotes.append(chart_config['notes'])
                try:
                    v.statusupdate(app, (msg + str(chart_config['notes'][0])[11:-1]), 2)
                    logging.info(msg + str(chart_config['notes'][0])[11:-1])
                except:  # If there's no worksheet values, this will happen
                    v.statusupdate(app, msg, 2)
                    logging.info(chart_config['notes'])


                if len(chart_config['bases']) > 0:
                    footertext.append(str(chart_config['bases'])[1:-1])
                if len(title_question) > 0:
                    chart_config['data question'], chart_config['chart title'] = er.longestval(title_question)
                    footertext.append(chart_config['data question'])
                if len(chart_config['note']) > 0:
                    footertext.append(str(chart_config['note'])[1:-1])
                df = wbdata[page][chart]['frame']
                chart_data = assignchartdata(df, chart_config)
                preflightlst.append(error_list)

                if intendedchart == 'TABLE':
                    for shape in shapeslist:
                        if 'TABLE' in shape:
                            if shape not in used_ph and chart_data not in used_data:
                                t_banding = chart_config['banding']
                                phidx = shapeslist[shape]['index']
                                placeholder = slide.placeholders[phidx]
                                chc.create_table(df, placeholder, chart_config)
                                if chart_config['*HEAT MAP'] is True:
                                    add_gradient_legend(slide)
                                used_ph.append(shape)
                                used_data.append(chart_data)
                elif intendedchart == 'STAT':
                    for shape in shapeslist:
                        if 'BODY' in shape:
                            if shape not in used_ph and chart_data not in used_data:
                                if shapeslist[shape]['height'] < 3 and shapeslist[shape]['width'] < 3:
                                    phidx = shapeslist[shape]['index']
                                    placeholder = slide.placeholders[phidx]
                                    text_frame = placeholder.text_frame
                                    insert_text(chart_data, text_frame)
                                    used_ph.append(shape)
                                    used_data.append(chart_data)
                elif intendedchart == 'PICTURE':
                    for shape in shapeslist:
                        if 'PICTURE' in shape:
                            if shape not in used_ph and chart_data not in used_data:
                                phidx = shapeslist[shape]['index']
                                placeholder = slide.placeholders[phidx]
                                picture = placeholder.insert_picture(chart_data)
                                used_ph.append(shape)
                                used_data.append(chart_data)
                else:
                    for shape in shapeslist:
                        if 'CHART' in shape:
                            if shape not in used_ph:
                                if chart_data not in used_data:
                                    phidx = shapeslist[shape]['index']
                                    placeholder = slide.placeholders[phidx]
                                    chc.create_chart(df, placeholder, chart_data, chart_config)
                                    used_ph.append(shape)
                                    used_data.append(chart_data)

            else:  # If it is page meta
                pagecontent = wbdata[page][chart]
                if len(pagecontent) > 0:
                    textlst = wbdata[page][chart]['page copy']
                    for shape in shapeslist:
                        if 'BODY' in shape:
                            if shape not in used_ph and textlst not in used_data:
                                if (shapeslist[shape]['height'] > 3) or (shapeslist[shape]['height'] > 2 and shapeslist[shape]['width'] > 9):
                                    phidx = shapeslist[shape]['index']
                                    placeholder = slide.placeholders[phidx]
                                    text_frame = placeholder.text_frame
                                    insert_text(textlst, text_frame)
                                    used_ph.append(shape)
                                    used_data.append(textlst)
                if len(pagecontent['callouts']) > 0:
                    shapes = slide.shapes
                    for callout in pagecontent['callouts']:
                        c = pagecontent['callouts'][callout]
                        width, height, left, top = Inches(c[0]), Inches(c[1]), Inches(c[2]), Inches(c[3])
                        shape = shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
                        shape.shadow.inherit = False
                        shape.line.fill.background()
                        fill = shape.fill
                        if c[4] == None:
                            fill.background()
                        elif c[4] == 'white':
                            fill.solid()
                            fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
                            fill.fore_color.brightness = 1

                        if c[6] == 'c':
                            alignment = PP_ALIGN.CENTER
                        else:
                            alignment = PP_ALIGN.LEFT
                        callout_formatting(shape, c[5], alignment, c[7], c[8])

        try:
            create_footer(slide, footertext, chart_config['directional check'])
        except UnboundLocalError:
            logging.info('d_check not defined')
        except AttributeError:
            logging.info('No Footer Placeholder found')

        notesinsert(slide, slidenotes)  # Add notes to every slide
        preflight(slide, preflightlst)  # Add boxes with error messages to applicable slides
    prs.save(trusave)
    v.statusupdate(app, ' File Saved', 2)