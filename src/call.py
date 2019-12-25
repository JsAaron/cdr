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
    layer = d1.getLayer("秒秒学装饰")
    # 必须设置活动的layer，这样调用vb.exe才会在这个layer的内部
    layer.Activate()
    imgShape = d1.addImage(layer,"C:\\Users\\Administrator\\Desktop\\111\\1.png")
    #设置单位像素
    d1.doc.Unit = 5
    ellipse = layer.CreateEllipse(100, 100, 500, 500)
    imgShape.AddToSelection()
    imgShape.AddToPowerClip(ellipse)


# 测试分组
def testGroup():
    cdrObj = CDR()
    layer = cdrObj.getLayer("秒秒学装饰")
    s1 =  layer.FindShape("test1")
    s2 =  layer.FindShape("test2")
    s3 =  layer.FindShape("test3")
    s4 =  layer.FindShape("test4")

    # 创建4个边界三角形
    if s1 == None:
        s1 = cdrObj.drawDecorationTriangle("test1",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   
    if s2 == None:
        s2 = cdrObj.drawDecorationTriangle("test2",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')   
    if s3 == None:
        s3 = cdrObj.drawDecorationTriangle("test3",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')  
    if s4 == None:
        s4 = cdrObj.drawDecorationTriangle("test4",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')   


# 测试组对象移动
def addShapeToGroup():
    cdrObj = CDR()
    layerObj = cdrObj.getLayer('秒秒学装饰')
    g1 = cdrObj.groupShapeObjs(layerObj,"占位组",)
    g2 = cdrObj.groupShapeObjs(layerObj,"子组占位组1",g1)
    g3 = cdrObj.groupShapeObjs(layerObj,"子组占位组2",g2)
    g4 = cdrObj.groupShapeObjs(layerObj,"子组占位组3",g3)

    s1 =  layerObj.FindShape("test1")
    if s1 == None:
       s1 = cdrObj.drawDecorationTriangle("test1",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   

    # 增加一个对象到组
    cdrObj.addShapeToGroup(g4,s1)


# 从组中删除一个对象，维持组的持久性
def removGroupShapeObjs():
    cdrObj = CDR()
    layerObj = cdrObj.getLayer('秒秒学装饰')
    g1 = cdrObj.groupShapeObjs(layerObj,"占位组",)
    g2 = cdrObj.groupShapeObjs(layerObj,"子组占位组1",g1)
    g3 = cdrObj.groupShapeObjs(layerObj,"子组占位组2",g2)
    g4 = cdrObj.groupShapeObjs(layerObj,"子组占位组3",g3)

    s1 =  layerObj.FindShape("test1")
    if s1 == None:
       s1 = cdrObj.drawDecorationTriangle("test1",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   

    # 增加一个对象到组
    cdrObj.addShapeToGroup(g4,s1)
    # 从组中删除一个对象，但是保持组的持久性
    cdrObj.removGroupShapeObjs(g4,s1)


# 从组中删除一个对象，不维持持组的持久性
def deleteGroupShapeObjs():
    cdrObj = CDR()
    layerObj = cdrObj.getLayer('秒秒学装饰')
    g1 = cdrObj.groupShapeObjs(layerObj,"占位组",)
    g2 = cdrObj.groupShapeObjs(layerObj,"子组占位组1",g1)
    g3 = cdrObj.groupShapeObjs(layerObj,"子组占位组2",g2)
    g4 = cdrObj.groupShapeObjs(layerObj,"子组占位组3",g3)

    s1 =  layerObj.FindShape("test1")
    if s1 == None:
       s1 = cdrObj.drawDecorationTriangle("test1",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   

    # 增加一个对象到组
    cdrObj.addShapeToGroup(g4,s1)
    cdrObj.deleteGroupShapeObjs(g4,s1)

# 修改文本
def modifyParaText():
    cdrObj = CDR()
    obj = cdrObj.insertParaText([0, 0, 120, 500],'测试1','我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111')
    cdrObj.modifyParaText(obj,'dfs123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111',[5, 50, 100, 50],'','')


def moveToMiddle():
    cdrObj = CDR()
    # cdrObj.moveToLandscapeMiddle("测试1")
    # cdrObj.moveToLeft('测试1')
    # cdrObj.moveToRight('测试1')
    # cdrObj.moveToTop('测试1')
    # cdrObj.moveToBottom('测试1')
    # cdrObj.moveToVerticalMiddle('测试1')
    cdrObj.moveToCenter('测试1')


if __name__ == '__main__':
    moveToMiddle()
    # modifyParaText()
    # cdrObj = CDR()
    # cdrObj.addFolder(1)
    # modifyParaText()
    # removGroupShapeObjs()
    # deleteGroupShapeObjs()
    # testPowerClip()
    # testAccessGroup()
    # testGroup()
    # drawDecorationTriangle()
    # drawDecorationTriangle()
    # drawDecorationTriangle()
    # togglePage()
    # print( test.get("aaaa") ==None)
    # getContent()
    # open()
    # setContent()
    

