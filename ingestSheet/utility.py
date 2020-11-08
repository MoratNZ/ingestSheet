def insertIntoDict(targetDict, headerArray, value):
    """Inserts a value into a dict, either directly, or via one or more child dicts. 

    If those child dicts don't exist, they're silently created.

    Args:
        targetDict (dict): The dict the value needs to be inserted into
        headerArray (list): An array of one or more strings, those strings being the keys / subkeys to associate the value with
        value (any): The value to be inserted into the dict

    Returns:
        dict: The dictt passed in, now with the value inserted into the appropriate place in its child hierachy
    """
    if targetDict is None:
        raise ValueError(
            "'None' passed to insertIntoDict as targetDict, instead of Dict")
    if headerArray is None:
        raise ValueError(
            "'None' passed to insertIntoDict as headerArray, instead of list")
    if len(headerArray) == 0:
        raise ValueError("Empty header array passed to insertIntoDict")

    elif len(headerArray) == 1:
        targetDict[headerArray[0]] = value

        return targetDict
    else:
        nextHeader = headerArray[-1]

        if nextHeader not in targetDict or not isinstance(targetDict[nextHeader], dict):
            targetDict[nextHeader] = {}

        nextDict = targetDict[nextHeader]

        targetDict[nextHeader] = insertIntoDict(
            nextDict, headerArray[:-1], value)

        return targetDict


def makeCamelCase(string):
    string = string.replace("_", " ")
    words = string.split(" ")

    if len(words) < 2:
        return string.lower()
    else:
        capitalisedWords = [word.title()
                            for word in words[1:]]
        capitalisedWords.insert(0, words[0].lower())

        return "".join(capitalisedWords)
