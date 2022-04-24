from openpyxl import load_workbook


def get_workbook(file_name, data_only=True):
    return load_workbook(filename=file_name, data_only=data_only)

