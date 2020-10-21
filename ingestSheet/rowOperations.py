from .utility import insertIntoDict


def parseRow(headerArray, rowarray):
    rowDict = {}

    for cellIndex in range(len(headerArray)):
        if headerArray[cellIndex] is not None and len(headerArray[cellIndex] > 0):
            insertIntoDict(
                rowDict, headerArray[cellIndex], rowarray[cellIndex].value)

    return rowDict
