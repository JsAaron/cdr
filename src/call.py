from cdr import CDR

# 打开文档
def open():
    print(CDR('C:\\Users\\Administrator\\Desktop\\11.cdr'))

# 获页面内容
# 传递页面搜索
# 不传，默认获取所有页面数据
def getContent(pageIndex=""):
    print('返回', CDR().get(pageIndex))


# 获取指定页面的全部内容字段
def setContent(pageIndex="",path=""):
    data = {'logo': {'pageIndex': 1, 'value': 'C:\\Users\\Administrator\\Desktop\\111\\1.png'}}
    CDR().set(data)


# 页面页面，传递不同页面的索引 1开始， 第一页1，第二页2
def togglePage():
    print(CDR().togglePage(2))
    

# 创建边界三角形
def drawDecorationTriangle():
    # CDR().groupDecorationTriangle()
    CDR().drawDecorationTriangle("test",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   
    CDR().drawDecorationTriangle("test",{"background-color":[255, 0, 0]},{"bottom":300,"right":600},'righttop')   
    CDR().drawDecorationTriangle("test",{"background-color":[255, 0, 0]},{"top":300,"left":600},'leftbottom')   
    CDR().drawDecorationTriangle("test",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')   


#测试图片裁剪
def testPowerClip():
    d1 = CDR()
    layer = d1.getShape("秒秒学装饰")
    # 必须设置活动的layer，这样调用vb.exe才会在这个layer的内部
    layer.Activate()
    imgShape = d1.addImage(layer,"C:\\Users\\Administrator\\Desktop\\111\\1.png")
    #设置单位像素
    d1.doc.Unit = 5
    ellipse = layer.CreateEllipse(100, 100, 500, 500)
    imgShape.AddToSelection()
    imgShape.AddToPowerClip(ellipse)


if __name__ == '__main__':
    testPowerClip()
    # drawDecorationTriangle()
    # drawDecorationTriangle()
    # drawDecorationTriangle()
    # togglePage()
    # print( test.get("aaaa") ==None)
    # getContent()
    # open()
    # setContent()
    

