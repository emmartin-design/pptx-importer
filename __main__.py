from appJar import gui

from pptx_handlers.template_reader import PPTXTemplate
from pptx_handlers.pptx_creator import create_report
from data_handlers.report_outlines import get_outline
from utilities.utility_functions import replace_chars, country_list, get_df_from_worksheet

version = 1.0


class AppJarTab:
    """
    This is a parent class for tabs
    It handles the stop/start commands and basic naming
    Child classes use "add_tab_content" for specific values
    """

    def __init__(self, parent, tab_name='Default Name'):
        self.parent = parent
        self.parent.app.startTab(tab_name)
        self.add_tab_contents()
        self.parent.app.stopTab()

    def add_tab_contents(self):
        pass


class AppJarLabelFrame:
    """
    This is a parent class for label frames
    It handles the stop/start commands and basic naming
    Child classes use "add_frame_content" for specific values
    """

    def __init__(
            self,
            parent,
            title=None,
            label=None,
            row=0,
            column=0,
            rowspan=0,
            colspan=0
    ):
        self.parent = parent
        self.parent.app.startLabelFrame(
            title,
            hideTitle=title is None,
            label=label,
            row=row,
            column=column,
            rowspan=rowspan,
            colspan=colspan
        )
        self.add_frame_content()
        self.parent.app.stopLabelFrame()

    def add_frame_content(self):
        pass


class DTVValuesLabelFrame(AppJarLabelFrame):
    """
    This app initiates and styles the DTV Values frame
    """

    def __init__(self, parent):
        super().__init__(
            parent,
            title='DTV Vals',
            label='DirecTV Segment Values',
            row=0,
            column=4,
            rowspan=2
        )

    def add_frame_content(self):
        for label_idx, entry_field in enumerate(['QSR', 'FC', 'MID', 'CD']):
            self.parent.app.addLabel(f'l{label_idx}', entry_field)
            self.parent.app.addEntry(entry_field)


class SettingsTab(AppJarTab):
    """
    This tab initiates and styles the settings tab
    """

    def __init__(self, parent):
        super().__init__(parent, tab_name='Settings')

    def add_tab_contents(self):
        self.parent.app.addCheckBox('Use Old Template', row=0, column=0)
        self.parent.app.setCheckBox('Use Old Template', ticked=True, callFunction=False)
        DTVValuesLabelFrame(self.parent)


class PowerPointImporterTab(AppJarTab):
    """
    This tab initiates and styles the Power Point Importer controls
    """
    default_template = 'templates/DATAIMPORT.pptx'
    old_template = 'templates/OLD_TEMPLATE.pptx'

    report_options = [
        'General Import',
        'Global Navigator Country Reports',
        'LTO Scorecard Report',
        'DirecTV Scorecard',
        'Quarterly Consumer KPIs',
    ]

    def __init__(self, parent):
        super().__init__(parent, tab_name='PowerPoint Importer')

    def add_tab_contents(self):
        self.parent.app.addLabelOptionBox('Report Type', self.report_options, row=0, column=2)
        self.parent.app.addEntry('xlsx_file', colspan=2, row=1, column=1)
        self.parent.app.addButton('Select File', self.select_xlsx, row=1, column=3)
        self.parent.app.addButton('Begin', self.press, row=2, column=1, colspan=3)
        self.parent.app.setButton('Begin', '     Begin     ')

    def select_xlsx(self):
        xslx_file = self.parent.app.openBox(
            title="Choose Import Data",
            dirName=None, fileTypes=[('excel worksheets', '*.xlsx'), ('excel worksheets', '*.xlsm')],
            parent=None,
            multiple=False,
            mode='r'
        )
        self.parent.app.setEntry('xlsx_file', xslx_file, callFunction=False)

    def create_filename(self, report):
        report = replace_chars(report, ('/', '_'), ('\\', '_'), ('_', '_'))
        file = self.parent.app.getEntry('xlsx_file')
        file = replace_chars(file, ('.xlsx', ''), ('.xlsm', ''))
        file = f"{file}{'' if report is None else f'_{report}'}.pptx"
        return file

    def get_entertainment_values(self):
        dtv_et_totals = [self.parent.app.getEntry(value) for value in ['QSR', 'FC', 'MID', 'CD']]
        all_are_filled_in = all([x != '' for x in dtv_et_totals])
        all_are_blank = all([x == '' for x in dtv_et_totals])
        assert all_are_filled_in or all_are_blank
        return dtv_et_totals

    def clear_status(self):
        for idx in range(1, 4):
            self.parent.app.setStatusbar('', field=idx)
            self.parent.app.setStatusbarBg("white", field=idx)

    def get_kpi_report_list(self):
        try:
            kpi_df = get_df_from_worksheet(self.parent.app.getEntry('xlsx_file'), 0)
            report_list = [x for x in kpi_df[kpi_df.columns[0]] if 'Avg' not in x]
            return report_list
        except TypeError:
            return [None]

    def get_report_list(self, report_type):
        parameters = {
            'General Import': [None],
            'Global Navigator Country Reports': country_list,
            'LTO Scorecard Report': [None],
            'DirecTV Scorecard': [None],
            'Quarterly Consumer KPIs': self.get_kpi_report_list()
        }
        return parameters.get(report_type)

    def get_verbatims(self, excel_file):
        if 'KPI' in self.parent.app.getOptionBox('Report Type'):
            self.parent.app.setStatusbar(f'Reading Verbatims', field=1)
            return get_df_from_worksheet(excel_file, 2)
        return None

    def press(self):
        """
        When the button is pressed, the following functions are called to create the report
        A new template and report_data class must be instituted for each new report in a list
        """

        report_type = self.parent.app.getOptionBox('Report Type')

        for report in self.get_report_list(report_type):
            print(f'Working on the {"" if report is None else report} report'.replace('  ', ' '))
            self.clear_status()
            self.parent.app.setStatusbar('Reading Template', field=0)
            use_old_template = self.parent.app.getCheckBox('Use Old Template')
            template_file = self.old_template if use_old_template else self.default_template
            template = PPTXTemplate(template_file, report_type)
            excel_file = self.parent.app.getEntry('xlsx_file')
            entertainment = self.get_entertainment_values()
            verbatims = self.get_verbatims(excel_file)
            self.parent.app.setStatusbarBg("gray", field=0)

            report_text = "" if report is None else f'{report} '
            self.parent.app.setStatusbar(f'Reading {report_text}Excel File', field=1)
            report_data = get_outline(
                report_type,
                excel_file,
                entertainment=entertainment,
                verbatims=verbatims,
                report_focus=report
            )
            self.parent.app.setStatusbarBg("gray", field=1)

            self.parent.app.setStatusbar(f'Creating {report_text}PPTX file', field=2)
            prs = create_report(report_data, template)
            self.parent.app.setStatusbarBg("gray", field=2)

            prs.save(self.create_filename(report))
            self.parent.app.setStatusbar('Report saved.', field=3)
            self.parent.app.setStatusbarBg("gray", field=3)
            print('Report Complete')


class MainApp:
    """
    This controls the overall structure of the app GUI
    Tabs and frames have separate classes initiated in this parent class
    """

    version = 2.0

    def __init__(self, main_app):
        self.app = main_app

        self.set_overall_style()
        self.set_title_bar()
        self.add_status_bar()
        self.add_tabbed_frames()

        self.pptx_tab = None
        self.xlsx_tab = None
        self.options_tab = None

        self.app.go()

    def set_overall_style(self):
        self.app.setTitle('PowerPoint Importer')
        self.app.setFont(14)
        self.app.setBg("white")
        self.app.setPadding([20, 20])
        self.app.setInPadding([20, 20])
        self.app.setStretch('both')

    def set_title_bar(self):
        title = f"PowerPoint Data Import\nVersion {self.version} Beta"
        self.app.addLabel('title', title, row=0, column=0)

    def add_status_bar(self):
        self.app.addStatusbar(fields=4)
        self.app.setStatusbarBg("white")
        self.app.setStatusbarFg("black")

    def add_tabbed_frames(self):
        app.startTabbedFrame("TabbedFrame", colspan=7, rowspan=1)
        app.setTabbedFrameBg('TabbedFrame', "white")
        app.setTabbedFrameTabExpand("TabbedFrame", expand=True)

        # Frame classes here
        self.pptx_tab = PowerPointImporterTab(self)
        self.xlsx_tab = None
        self.options_tab = SettingsTab(self)

        app.stopTabbedFrame()


if __name__ == '__main__':
    """
    This is what starts the app
    """
    with gui('File Selection', '1000x600') as app:
        MainApp(app)
