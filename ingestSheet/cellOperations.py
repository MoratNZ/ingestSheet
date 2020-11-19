import openpyxl
import re

whitespaceRe = re.compile('^\s*$')


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


def isChildMergedCell(cell):
    if not isinstance(cell, openpyxl.cell.MergedCell):
        return False
    else:
        if cell.value is None:
            return True
        else:
            return False


def getCellValue(cell):
    if isChildMergedCell(cell):
        column = cell.column
        if column == 1:
            return None
        else:
            return getCellValue(cell.parent.cell(row=cell.row, column=(cell.column - 1)))
    else:
        return cell.value
