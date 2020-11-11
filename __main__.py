from appJar import gui
from datetime import datetime, date, time
from pptx import Presentation
import os
import glob
import logging
import sys
import re
import pandas as pd
import numpy as np
from datetime import datetime, date, time

# Internal module pieces
import variables as v
import templatereader as tr
import slidecreator as sc
import excelreader as er
import excelreader2 as er2


# To package as EXE run "python -m PyInstaller --onefile --windowed __main__.py" in terminal


version = v.version
reporttypes = v.reporttypes

# Creates a new log for each day the software is run
# This makes it more accessible, considering the low volume
# If volume increases, may switch to one file, and use fileneame in log.
loggingfilename = 'gen_' + datetime.now().strftime('%y%B%d') + '.log'
log_dir = os.path.join(os.path.normpath(os.getcwd()), 'logs')
logfn = os.path.join(log_dir, loggingfilename)
logging.basicConfig(filename=logfn, level=logging.DEBUG, format='%(lineno)d:%(levelname)s:%(message)s')

# Allows only XLSX files to be selected
def xlsxselect():
    xslxfile = app.openBox(title="Choose Import Data", dirName=None, fileTypes=[('excel worksheets', '*.xlsx')],
                           parent=None, multiple=False, mode='r')
    app.setEntry('xlsxfile', xslxfile, callFunction=False)


def templateselect():
    template = app.openBox(title="Choose Different Template", dirName=None, fileTypes=[('PowerPoint files', '*.pptx')],
                           parent=None, multiple=False, mode='r')
    app.setEntry('template', template, callFunction=False)

def clearstatus():
    msg = ""
    for idx in range(1,4,1):
        app.setStatusbar(msg, field=idx)
        app.setStatusbarBg("white", field=idx)


# Defines what happens when the import button is pressed.
def press():
    clearstatus()
    datafile = app.getEntry('xlsxfile')
    type = app.getOptionBox('Report Type')
    template = app.getEntry('template')
    v.statusupdate(app, 'Loading Template', 0)
    if type in ['General Import', 'Global Navigator Country Reports']:
        templatedata = tr.readtemplate(template, 'general')
        v.statusupdate(app, 'Opening Excel', 1)
        if type in ['General Import', 'Global Navigator Country Reports']:
            try:
                if 'Global' in type:
                    reportlist = v.countrylist
                else:
                    reportlist = [None]
                dsave = app.directoryBox(title='Where should report(s) be saved?', dirName=None, parent=None)
                for r in reportlist:
                    if r == None:
                        reportname = 'dataimport | ' + datetime.now().strftime("%B %d, %Y")
                        fulld = datafile[:-5] + '.pptx'
                        msg = 'Processing Data: '
                    else:
                        reportname = r + 'Country Report'
                        pptxname = r + '.pptx'
                        msg = r + ' report: '
                        fulld = dsave + "/" + pptxname
                    if app.getCheckBox('Use Single-Color Import'):
                        wbdata = er2.readbook(app, datafile, reportname, country=r)
                    else:
                        wbdata = er.readbook(app, datafile, reportname, country=r)
                    sc.data_import(app, template, wbdata, templatedata, fulld, msg=msg)
                    logging.info(fulld + ' saved.')
                    logging.info('Import Complete')
                v.statusupdate(app, 'Complete', 3)
            # except UnboundLocalError:
                # v.errorupdate(app, 'UnboundLocalError')
            except PermissionError:
                v.errorupdate(app, 'PermissionError')
            # except IndexError:
                # v.errorupdate(app, 'IndexError')

    elif type in ['LTO Scorecard Report', 'Value Scorecard Report']:
        valrpt = type == 'Value Scorecard Report'
        try:
            templatedata = tr.readtemplate(template, 'scorecard')
            v.statusupdate(app, 'Reading Data', 1)
            wbdata = er.scorecard(datafile, valrpt)
            v.statusupdate(app, 'Importing Data', 2)
            fulld = datafile[:-5] + '.pptx'
            msg = 'Processing Data: '
            sc.data_import(app, template, wbdata, templatedata, fulld, msg=msg)
            v.statusupdate(app, 'Complete', 3)
        # except UnboundLocalError:
            # v.errorupdate(app, 'UnboundLocalError')
        except PermissionError:
            v.errorupdate(app, 'PermissionError')
        # except ValueError:
            # v.errorupdate(app, 'ValueError')
    elif type == 'DirecTV Scorecard':
        templatedata = tr.readtemplate(template, 'dtv')
        v.statusupdate(app, 'Reading Data', 1)
        wbdata = er.dtv_reader(datafile)
        v.statusupdate(app, 'Importing Data', 2)
        fulld = datafile[:-5] + '.pptx'
        msg = 'Processing Data: '
        sc.data_import(app, template, wbdata, templatedata, fulld, msg=msg)
        v.statusupdate(app, 'Complete', 3)

with gui('File Selection', '1000x600') as app:
    print('Starting Software, please wait...')
    app.setTitle('PowerPoint Importer')
    app.setFont(14)
    app.setBg("white")
    app.setPadding([20, 20])
    app.setInPadding([20, 20])
    app.setStretch("both")

    app.addLabel('title', ('PowerPoint Data Import \n' + version), row=0, column=0)
    app.addWebLink('View Instructions', "https://cspnet1-my.sharepoint.com/:w:/g/personal/emartin_technomic_com/Eag8lFMxFTlApMXiLQwYudIBDc5CTE3TIWmR6TLr2a41Qg?e=02wf2i", row=0, column=1)

    app.addStatusbar(fields=4)
    app.setStatusbarBg("white")
    app.setStatusbarFg("black")

    app.startTabbedFrame("TabbedFrame", colspan=7, rowspan=1)
    app.setTabbedFrameBg('TabbedFrame', "white")
    app.setTabbedFrameTabExpand("TabbedFrame", expand=True)

    app.startTab('PowerPoint Importer')
    app.addLabelOptionBox('Report Type', reporttypes, row=0, column=2)
    app.addEntry('xlsxfile', colspan=2, row=1, column=1)
    app.addButton('Select File', xlsxselect, row=1, column=3)
    app.addButton('Begin', press, row=2, column=1, colspan=3)
    app.setButton('Begin', '     Begin     ')
    app.stopTab()

    app.startTab('Settings')
    app.addCheckBox('Use Single-Color Import')
    app.addEntry('template')
    app.setEntry('template', v.defaulttemplate, callFunction=False)
    app.addButton('Select Alternate Template', templateselect)
    app.stopTab()

    app.stopTabbedFrame()

    app.go()
