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


def getCellValue(cell):
    if isChildMergedCell(cell):
        return getCellValue(cell.parent.cell(row=cell.row, column=(cell.column - 1)))
    else:
        return cell.value
