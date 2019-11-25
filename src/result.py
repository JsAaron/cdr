state = False
totalpages = 0
fnreturn = {}
inputData = {}
inputFiled = {}

# 保存input数组显示的字段
def saveInputFiled(key):
    if key == "url":
        inputFiled["bjnews"] = "url+bjnews"
    elif key == "bjnews":
        inputFiled["url"] = "url+bjnews"
    elif key == "mobile":
        inputFiled["phone"] = "mobile+phone"
    elif key == "phone":
        inputFiled["mobile"] = "mobile+phone"
    elif key == "email":
        inputFiled["qq"] = "email+qq"
    elif key == "qq":
        inputFiled["email"] = "email+qq"

    inputFiled[key] = True


def saveData(pageIndex, key, value, overflow=False):
    json = {}

    # 溢出了
    if overflow == True:
        json["overflow"] = True

    json["pageIndex"] = pageIndex
    json["value"] = value
    inputData[key] = json


# 通过段落去匹配出key来
def valueTokey(pageIndex, name, tempShape):
    p = tempShape.Text.Story.Paragraphs
    v1 = p.Item(1).Text
    v2 = p.Item(2).Text

    # '电话手机一组
    if name == "mobile" or name == "phone":
        saveData(pageIndex, "mobile", v1)
        saveData(pageIndex, "phone", v2)

    if name == "email" or name == "qq":
        saveData(pageIndex, "email", v1)
        saveData(pageIndex, "qq", v2)

    if name == "url" or name == "bjnews":
        saveData(pageIndex, "url", v1)
        saveData(pageIndex, "bjnews", v2)


# 填充默认值给外部
def fillDefault(pageIndex, key, tempShape):
    # 保存当前值
    saveData(pageIndex, key, tempShape.Text.Story.Text)
    # 填充默认值
    if key == "url":
        saveData(pageIndex, "bjnews", "")
    elif key == "bjnews":
        saveData(pageIndex, "url", "")
    elif key == "mobile":
        saveData(pageIndex, "phone", "")
    elif key == "phone":
        saveData(pageIndex, "mobile", "")
    elif key == "email":
        saveData(pageIndex, "qq", "")
    elif key == "qq":
        saveData(pageIndex, "email", "")


# 保存获取的值
# 可能有分组组合的情况，所以需要找到字段合计，然后找到分组的数组
def saveValue(pageIndex, key, tempShape, determine, onlyFill):

    # 去重
    if inputData.get(key):
        #print("重复保存",key)
        return

    # 如果只是填充默认值,仅针对图片的读
    if onlyFill:
        saveData(pageIndex, key, "")
        # print("填充默认图片", key)
        return

    # 是否存在需要分解的数据
    hasRange = determine.getRangeScope(key)
    if hasRange:
        # 一个字段有上下2行,可能是被改变过，需要分解
        if tempShape.Text.Story.Paragraphs.Count == 2:
            valueTokey(pageIndex, key, tempShape)
        else:
            # 填充默认值
            fillDefault(pageIndex, key, tempShape)
    else:
        # 直接保存
        saveData(pageIndex, key, tempShape.Text.Story.Text,
                 tempShape.Text.Overflow)


# 页码总数
def setPageTotal(count):
    global totalpages 
    totalpages = count

# 生成返回数据
def retrunData():
    return {
        "totalpages":totalpages,
        "fileds": inputFiled,
        "data": inputData
    }
