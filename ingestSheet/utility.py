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
