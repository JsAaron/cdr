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


# 替换图片
def __replaceImage(doc, tempShape, key, typeName):
    # 中心点
    doc.ReferencePoint = 9
    centerX = tempShape.CenterX
    centerY = tempShape.CenterY
    SizeWidth = tempShape.SizeWidth
    SizeHeight = tempShape.SizeHeight

    # 返回或设置形状所在的图层
    parentLayer = tempShape.Layer
    imageType = 802
    imagePath = prarm.getExternalValue(key)
    parentLayer.Activate()

    # jpg类型
    args = imagePath.split(".jpg")

    if args.Count == 2:
        imageType = 774

    # 修改图片必须是显示状态才可以
    fixVisible = False
    if parentLayer.Visible == False:
        fixVisible = True
        parentLayer.Visible = True

    parentLayer.Import(imagePath, imageType)
    # 重新设置图片
    dfShapes = doc.Selection.Shapes
    # 插入成功才删除图片
    if dfShapes.Count > 0:
        for item in dfShapes:
            item.Name = typeName
            item.SetSize(SizeWidth, SizeHeight)
            item.SetPositionEx(9, centerX, centerY)
    tempShape.Delete()

    # 如果修改了图片状态
    if fixVisible:
        parentLayer.Visible = False


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


def accessImage(doc, allShapes):
    # '递归检测形状,并替换图片
    for tempShape in allShapes:
        cdrGroupShape = 7
        cdrBitmapShape = 5

        # 组
        if tempShape.Type == cdrGroupShape:
            accessImage(doc, tempShape.Shapes)

        if tempShape.Type == cdrBitmapShape:
            if tempShape.Name == "二维码" and prarm.hasValue("qrcode"):
                __replaceImage(doc, tempShape, "qrcode", "二维码")
            elif tempShape.Name == "Logo" and prarm.hasValue("logo"):
                __replaceImage(doc, tempShape, "logo", "Logo")
            elif tempShape.Name == "Logo2" and prarm.hasValue("logo2"):
                __replaceImage(doc, tempShape, "logo2", "Logo2")
