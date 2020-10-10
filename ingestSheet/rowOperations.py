from .utility import insertIntoDict


def parseRow(headerArray, rowarray):
    rowDict = {}

    for cellIndex in range(len(headerArray)):
        if headerArray[cellIndex] is not None:
            insertIntoDict(
                rowDict, headerArray[cellIndex], rowarray[cellIndex].value)

    return rowDict
