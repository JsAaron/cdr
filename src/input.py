import prarm
import utils
import result
import prarm

# 图片的读，创建基本结构


def getImage(tempShape, pageIndex, determine):
    key = utils.getKeyEnglish(tempShape.Name)
    if key:
        result.saveValue(pageIndex, key, tempShape, determine, True)


# 获取文本
def getText(tempShape, pageIndex, determine):
    if tempShape.Text.Story.Text:
        key = utils.getKeyEnglish(tempShape.Name)
        if key:
            result.saveValue(pageIndex, key, tempShape, determine, False)


# 设置文本
def setText(tempShape, pageIndex, determine):
    key = utils.getKeyEnglish(tempShape.Name)
    value = ""
    if key:
        # 是否是处理范围
        if determine.getRangeScope(key):
            # 字段合并处理
            value = determine.getMergeValue(key)
        else:
            # 单独字段
            value = prarm.getExternalValue(key)

        # 如果有值
        if value:
            # 值不相等,替换
            if tempShape.Text.Story.Text != value:
                tempShape.Text.Story.Delete()
                tempShape.Text.Story.Replace(value)
        else:
            tempShape.Text.Story.Delete()


# 递归检测形状
def accessShape(doc, allShapes, determine, pageIndex):
    for tempShape in allShapes:
        cdrTextShape = 6
        cdrGroupShape = 7
        cdrBitmapShape = 5

        # 组
        if tempShape.Type == cdrGroupShape:
            accessShape(doc, tempShape.Shapes, determine, pageIndex)

        # 图片的读
        if tempShape.Type == cdrBitmapShape:
            if prarm.cmdCommand == "get:text":
                getImage(tempShape, pageIndex, determine)

        # 文字读写
        if tempShape.Type == cdrTextShape:
            # 读数据
            if prarm.cmdCommand == "get:text":
                getText(tempShape, pageIndex, determine)

            # 写数据
            if prarm.cmdCommand == "set:text":
                setText(tempShape, pageIndex, determine)
