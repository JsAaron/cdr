
cmdCommand = "get:text"
cmdExternalData = {}

def hasValue(key):
    # 如果没数据
    if len(cmdExternalData) == 0:
        return False

    # 如果有值
    if cmdExternalData.get(key) and cmdExternalData[key]["value"]:
        return True

    return False
