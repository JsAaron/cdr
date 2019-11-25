
cmdCommand = "get:text"
cmdExternalData = {}

# 探测是是否存在值


def hasValue(key):
    # 如果没数据
    if len(cmdExternalData) == 0:
        return False
    #如果有值
    if cmdExternalData[key] and cmdExternalData[key]["value"]:
        return True
    return False
