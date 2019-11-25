import globalData
import utils
import result

# 图片的读，创建基本结构
def getImage(tempShape, pageIndex, determine):
    key = utils.getKeyEnglish(tempShape.Name)
    if key:
        result.saveValue(pageIndex, key, tempShape, determine, True)


# 获取文本
def getText(tempShape, pageIndex, determine):
    print(123,tempShape.Text.Story.Text)
    if tempShape.Text.Story.Text:
        key = utils.getKeyEnglish(tempShape.Name)
        if key:
            result.saveValue(pageIndex, key, tempShape, determine, False)
        else:
            print("找不到对应的命名:"+tempShape.Name)



# 设置文本
def setText(tempShape, pageIndex, determine):
    key = utils.getKeyEnglish(tempShape.Name)
    value = ""
    # if key:
    #     #是否是处理范围
    #     if determine.getRangeScope(key):
    #         # 字段合并处理
    #         value = determine.getMergeValue(key)
    #     else:
    #         # 单独字段
    #         value = Param.getExternalValue(key)
            
        


# 递归检测形状
def accesstShape(doc, allShapes, determine, pageIndex):
    for tempShape in allShapes:
        cdrTextShape = 6
        cdrGroupShape = 7
        cdrBitmapShape = 5

        # 组
        if tempShape.Type == cdrGroupShape:
            accesstShape(doc, tempShape.Shapes, determine, pageIndex)

        # 图片的读
        if tempShape.Type == cdrBitmapShape:
            if globalData.cmdCommand == "get:text":
                getImage(tempShape, pageIndex, determine)

        # 文字读写
        if tempShape.Type == cdrTextShape:
            #读数据
            if globalData.cmdCommand == "get:text":
                getText(tempShape, pageIndex, determine)

            # # 写数据
            # if globalData.cmdCommand == "set:text":
            #     setText(tempShape, pageIndex, determine)