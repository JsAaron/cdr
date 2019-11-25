
cmdCommand = "get:text"
cmdExternalData = {}


def setCommand(value):
    global cmdCommand
    cmdCommand = value


def getCommand():
    return cmdCommand


def setExternalData(data):
    global cmdExternalData
    cmdExternalData = data


def getExternalValue(key):
    # 必须有值
    if len(cmdExternalData) == 0:
        return ""

    # 必须有属性
    if cmdExternalData.get(key):
        return cmdExternalData[key]["value"]

    return ""


def hasValue(key):
    # 如果没数据
    if len(cmdExternalData) == 0:
        return False

    # 如果有值
    if cmdExternalData.get(key) and cmdExternalData[key]["value"]:
        return True

    return False
