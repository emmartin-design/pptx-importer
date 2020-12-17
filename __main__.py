# To package as EXE run "python -m PyInstaller --onefile --windowed __main__.py" in terminal
from appJar import gui
from datetime import datetime

# Internal module pieces
import variables as v
import templatereader as tr
import slidecreator as sc
import excelreader as er
import excelreader2 as er2

version = v.version
report_config = v.report_config


# Allows only XLSX files to be selected
def xlsxselect():
    xslxfile = app.openBox(title="Choose Import Data", dirName=None, fileTypes=[('excel worksheets', '*.xlsx')],
                           parent=None, multiple=False, mode='r')
    app.setEntry('xlsxfile', xslxfile, callFunction=False)


def templateselect():
    template = app.openBox(title="Choose Different Template", dirName=None, fileTypes=[('PowerPoint files', '*.pptx')],
                           parent=None, multiple=False, mode='r')
    app.setEntry('template', template, callFunction=False)


def clear_status():
    for idx in range(1, 4, 1):
        app.setStatusbar('', field=idx)
        app.setStatusbarBg("white", field=idx)


def create_filename():
    file = app.getEntry('xlsxfile')
    filename = file.replace('.xlsx', '')
    v.log_entry('***********************************************')
    v.log_entry(('Start of import for ' + filename))
    return file, filename


# Defines what happens when the import button is pressed.
def press():
    clear_status()
    file, filename = create_filename()

    # Find the report type
    report_selection = app.getOptionBox('Report Type')
    report_list = report_config[report_selection]['report list']
    report_type = report_config[report_selection]['report type']
    report_suffix = report_config[report_selection]['report suffix']

    # Read Optional Value inputs
    dtv_et_totals = []
    for value in ['QSR', 'FC', 'MID', 'CD']:
        if app.getEntry(value) != '':
            dtv_et_totals.append(app.getEntry(value))
    if len(dtv_et_totals) < 4:
        if len(dtv_et_totals) > 0:
            v.log_entry('Must have 0 or 4 DTV values', level='warning', app_holder=app, fieldno=0)
            return
        else:
            dtv_et_totals = None


    # Find and read the template
    template = app.getEntry('template')
    v.log_entry('Loading Template', app_holder=app, fieldno=0)
    template_data = tr.collect_template_data(template, report_type)

    # Begin excel reading and report building process
    v.log_entry('Opening Excel', app_holder=app, fieldno=1)
    dsave = app.directoryBox(title='Where should report(s) be saved?', dirName=None, parent=None)

    for report in report_list:  # Allows for multiple reports from one excel
        if report is None:
            report_name = report_type + '_' + datetime.now().strftime("%B %d, %Y")
            full_directory = filename + '.pptx'
        else:
            report_name = report + report_suffix
            full_directory = dsave + '/' + report + '.pptx'

        # Find and read the excel file
        try:
            if report_type in ['general', 'global']:
                if app.getCheckBox('Use Single-Color Import'):  # statement will be removed after new selection method
                    wbdata = er2.data_collector(app, file, report_name, country=report)
                else:
                    wbdata = er.readbook(app, file, report_name, country=report)

            elif report_type in ['lto', 'value']:
                wbdata = er.scorecard(file, report_config[report_selection])

            elif report_type == 'dtv':
                wbdata = er.dtv_reader(file, dtv_et_totals)

            # Create pptx and drop in data
            sc.data_import(app, template, wbdata, template_data, full_directory, msg='Processing Data')
            v.log_entry(full_directory + ' saved.')
            v.log_entry('Import Complete', app_holder=app, fieldno=3)

        except PermissionError:
            v.log_entry('File open cannot save', level='warning', app_holder=app, fieldno=3)
        '''except ValueError:
            v.log_entry('Check Report Type', level='warning', app_holder=app, fieldno=3)
        except IndexError:
            v.log_entry('Data Selection Error', level='warning', app_holder=app, fieldno=3)
        except UnboundLocalError:
            v.log_entry('Check Report Type', level='warning', app_holder=app, fieldno=3)'''


with gui('File Selection', '1000x600') as app:
    # Assign app object to variables to hold for other modules
    v.app = app
    print('Starting Software, please wait...')
    app.setTitle('PowerPoint Importer')
    app.setFont(14)
    app.setBg("white")
    app.setPadding([20, 20])
    app.setInPadding([20, 20])
    app.setStretch('both')

    app.addLabel('title', ('PowerPoint Data Import \n' + version), row=0, column=0)
    app.addWebLink('View Instructions', "https://cspnet1-my.sharepoint.com/:w:/g/personal/emartin_technomic_com/Eag8lFMxFTlApMXiLQwYudIBDc5CTE3TIWmR6TLr2a41Qg?e=02wf2i", row=0, column=1)

    app.addStatusbar(fields=4)
    app.setStatusbarBg("white")
    app.setStatusbarFg("black")

    app.startTabbedFrame("TabbedFrame", colspan=7, rowspan=1)
    app.setTabbedFrameBg('TabbedFrame', "white")
    app.setTabbedFrameTabExpand("TabbedFrame", expand=True)

    app.startTab('PowerPoint Importer')
    app.addLabelOptionBox('Report Type', report_config.keys(), row=0, column=2)
    app.addEntry('xlsxfile', colspan=2, row=1, column=1)
    app.addButton('Select File', xlsxselect, row=1, column=3)
    app.addButton('Begin', press, row=2, column=1, colspan=3)
    app.setButton('Begin', '     Begin     ')
    app.stopTab()

    app.startTab('Settings')
    app.addCheckBox('Use Single-Color Import', row=0, column=0)
    app.addEntry('template', row=1, column=0, colspan=2)
    app.setEntry('template', v.default_template, callFunction=False)
    app.addButton('Select Alternate Template', templateselect, row=1, column=3)

    app.startLabelFrame('DTV Vals', hideTitle=False, label='DirecTV Segment Values (0.0)', row=0, column=4, rowspan=2)
    app.addLabel("l1", 'QSR')
    app.addEntry('QSR')
    app.addLabel("l2", 'FC')
    app.addEntry('FC')
    app.addLabel("l3", 'MID')
    app.addEntry('MID')
    app.addLabel("l4", 'CD')
    app.addEntry('CD')
    app.stopLabelFrame()
    app.stopTab()

    app.stopTabbedFrame()

    app.go()
