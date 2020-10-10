import openpyxl
import re
import json

whitespaceRe = re.compile('^\s*$')


def openWorkbook(workbookName):
    return openpyxl.load_workbook(filename=workbookName, read_only=True, data_only=True)


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


def insertIntoDict(targetDict, headerArray, value):
    if len(headerArray) == 1:
        targetDict[headerArray[0]] = value

        return targetDict
    else:
        nextHeader = headerArray[-1]

        if nextHeader not in targetDict:
            targetDict[nextHeader] = {}

        nextDict = targetDict[nextHeader]

        targetDict[nextHeader] = insertIntoDict(
            nextDict, headerArray[:-1], value)

        return targetDict


def parseRow(headerArray, rowarray):
    rowDict = {}

    for cellIndex in range(len(headerArray)):
        if headerArray[cellIndex] is not None:
            insertIntoDict(
                rowDict, headerArray[cellIndex], rowarray[cellIndex].value)

    return rowDict


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
