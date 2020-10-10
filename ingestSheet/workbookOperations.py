import openpyxl

def openWorkbook(workbookName):
    return openpyxl.load_workbook(filename=workbookName, read_only=True, data_only=True)