import prarm
import utils
import result
import json
import prarm
import subprocess
import urllib.parse

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
    print(key)
    imagePath = prarm.getExternalValue(key)
    data = "{'"+ key +"':{'pageIndex':'1','value':'"+ urllib.parse.quote(imagePath) +"'}}"
    cmdStr = ["D:\\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe", 'set:image', data]
    child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,stdin=subprocess.PIPE, stderr=subprocess.PIPE)
    for line in child.stdout.readlines():
        output = line.decode('UTF-8')
        print("output",output)

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
