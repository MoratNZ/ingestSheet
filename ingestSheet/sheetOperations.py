import openpyxl

import json

from .cellOperations import isEmptyCell, isChildMergedCell, getCellValue
from .rowOperations import parseRow


def parseHeaders(sheet, headerRowCount=1, headerColumnCount=1, maxColumnGap=1):
    currentColumn = headerColumnCount + 1
    blankColumns = 0

    returnArray = []

    for i in range(headerColumnCount):
        returnArray.append(None)

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


def parseSheet(sheet, headerRowCount=1, headerColumnCount=1, maxColumnGap=1):
    result = {}

    headers = parseHeaders(sheet, headerRowCount,
                           headerColumnCount, maxColumnGap)

    for rowIndex in range(headerRowCount + 1, sheet.max_row):
        labelCell = sheet.cell(row=rowIndex, column=headerColumnCount)

        if isEmptyCell(labelCell):
            next

        rowLabel = labelCell.value

        rowDict = parseRow(headers, sheet[rowIndex])

        result[rowLabel] = rowDict

    return result
