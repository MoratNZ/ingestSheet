import openpyxl
import re

whitespaceRe = re.compile('^\s*$')


def openWorkbook(workbookName):
    return openpyxl.load_workbook(filename=workbookName, read_only=True, data_only=True)


def parseHeaders(sheet, headerRowCount=1, headerColumnCount=1, maxColumnGap=1):
    currentColumn = headerColumnCount + 1
    blankColumns = 0

    returnArray = []

    for i in range(headerColumnCount):
        returnArray.append([])

    while True:
        columnArray = []
        for currentRow in range(headerRowCount, 0, -1):
            currentCell = sheet.cell(row=currentRow, column=currentColumn)
            if not isEmptyCell(currentCell):
                columnArray.append(getCellValue(currentCell))

        if len(columnArray) == 0:
            blankColumns += 1

            if blankColumns > maxColumnGap:
                break

        currentColumn += 1
        returnArray.append(columnArray)

    for i in range(maxColumnGap):
        returnArray.pop()

    return returnArray


def isChildMergedCell(cell):
    if isinstance(cell, openpyxl.cell.read_only.EmptyCell):
        return False
    elif isinstance(cell, openpyxl.cell.read_only.ReadOnlyCell):
        if cell.value is None:
            return True
        else:
            return False
    else:
        raise Exception(
            "Code has reached a state it shouldn't be able to. Cell has type: {}".format(type(cell)))


def isEmptyCell(cell):
    if isinstance(cell, openpyxl.cell.read_only.EmptyCell):
        return True
    else:
        cellValue = getCellValue(cell)

        if cellValue is None:
            return True
        elif isinstance(cellValue, str):
            if whitespaceRe.match(cellValue) is None:
                return False
            else:
                return True
        else:
            return False


def getCellValue(cell):
    if isChildMergedCell(cell):
        return getCellValue(cell.parent.cell(row=cell.row, column=(cell.column - 1)))
    else:
        return cell.value
