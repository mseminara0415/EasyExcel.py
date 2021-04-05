import win32com.client as win32
import os
import re
import sys
import shutil


class EasyExcel:
    """
    Excel Editor is designed to reduce the clutter of formatting excel files with
    Pywin32.
    """

    class ExcelCharacterLimit(Exception):
        def __init__(self, message):
            self.message = message

    class ExcludedCharacters(Exception):
        def __init__(self, message):
            self.message = message

    @staticmethod
    def initialize_excel(visible: bool, display_alerts: bool, screen_updating: bool, enable_events: bool):
        """
        Iniitialize Excel Object. User can set the parameters to determine if the process is visible or not.
        :param visible:
        :param display_alerts:
        :param screen_updating:
        :param enable_events:
        :return:
        """

        # Set up Excel Workbook
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = visible
            excel.Application.DisplayAlerts = display_alerts
            excel.ScreenUpdating = screen_updating
            excel.EnableEvents = enable_events
        except AttributeError:
            # Remove cache and try again.
            MODULE_LIST = [m.__name__ for m in sys.modules.values()]
            for module in MODULE_LIST:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(
                os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = visible
            excel.Application.DisplayAlerts = display_alerts
            excel.ScreenUpdating = screen_updating
            excel.EnableEvents = enable_events

        return excel

    def __init__(self, file_path: str):
        self.filepath = file_path
        self.excel = self.initialize_excel(False, False, False, False)
        self.wb = self.excel.Workbooks.Open(self.filepath)
        self.sheets = self.wb.Sheets

    @property
    def sheet_list(self):
        """
        Return a list of sheet names in workbook.
        :return:
        """
        return [sheet.Name for sheet in self.sheets]

    def color_scale(self, worksheet: str, cell_range_start=None,
                    cell_range_end=None, save=False):

        sheet = self.wb.Worksheets(worksheet)

        # Create Working Cell Range
        if cell_range_end is not None:
            working_range = sheet.Range(
                f"{cell_range_start}:{cell_range_end}")
        else:
            working_range = sheet.Range(f"{cell_range_start}")

        working_range.FormatConditions.AddColorScale(ColorScaleType=3)

        if save:
            self.save()

    def merge_cells(self, worksheet, cell_range_start=None, cell_range_end=None,
                    center_text=True, save=False):
        """
        Merge Excel Cell Range. Also Provides an option to center align.
        :return:
        """
        sheet = self.wb.Worksheets(worksheet)

        # Create Working Cell Range
        if cell_range_end is not None:
            working_range = sheet.Range(
                f"{cell_range_start}:{cell_range_end}")
        else:
            working_range = sheet.Range(f"{cell_range_start}")

        working_range.MergeCells = True

        # Center Text
        if center_text:
            sheet.Range(f"{cell_range_start}").HorizontalAlignment = -4108

        if save:
            self.save()

    def bold_cells(self, worksheet, cell_range_start=None, cell_range_end=None,save=False):
        """
        Bold cell(s) within specified worksheet.
        :param worksheet:
        :param cell_range_start:
        :param cell_range_end:
        :param save:
        :return:
        """

        sheet = self.wb.Worksheets(worksheet)

        # Create Working Cell Range
        if cell_range_end is not None:
            working_range = sheet.Range(
                f"{cell_range_start}:{cell_range_end}")
        else:
            working_range = sheet.Range(f"{cell_range_start}")

        working_range.Font.Bold = True

        if save:
            self.save()


    def add_sheet(self, name, save=False):
        """
        add worksheet to workbook.
        :param name:
        :param save:
        :return:
        """
        wb = self.wb
        ws = wb.Worksheets.Add()

        excluded_characters = [r"\/?*,[,]:."]
        matched_list = [characters in excluded_characters for characters in name]
        string_name_includes_excluded = all(matched_list)

        if string_name_includes_excluded:
            raise EasyExcel.ExcludedCharacters(f"Worksheet names cannot include {excluded_characters}")

        elif len(name) > 31:
            raise EasyExcel.ExcelCharacterLimit("Please Limit name to 31 characters")

        elif name in self.sheet_list:
            number_of_sheets = len(self.sheet_list) + 1
            ws.Name = name + str(number_of_sheets)

        else:
            ws.Name = name

        if save:
            self.save_work()

    def save(self):
        """

        :return:
        """
        self.wb.Save()

    def close_workbook(self):
        """
        Close and Save workbook.
        :return:
        """
        self.wb.Close(True)
        self.wb = None
        self.excel = None
