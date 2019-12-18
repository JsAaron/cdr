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


    # # 创建一个组对象
    # cdrObj.groupShape(layer,"组1",['test1','test2'])

    # 在秒秒学结构层下 创建占位组名占位组
    # cdrObj.accessGroup("占位组",'秒秒学结构')


    # 往组对象，添加2个新的对象
    # cdrObj.addShapeToGroup(newGroups,['test1','test4'])



# 测试组对象占位
def testAccessGroup():
    cdrObj = CDR()
    layerObj = cdrObj.getLayer('秒秒学装饰')
    g1 = cdrObj.accessGroup("占位组",layerObj)
    g2 = cdrObj.accessGroup("子组占位组1",g1,layerObj)
    # cdrObj.accessGroup("子组占位组2",g2,layerObj)
    # cdrObj.accessGroup("子组占位组2",g2,layerObj)


if __name__ == '__main__':
    # testPowerClip()
    testAccessGroup()
    # testGroup()
    # drawDecorationTriangle()
    # drawDecorationTriangle()
    # drawDecorationTriangle()
    # togglePage()
    # print( test.get("aaaa") ==None)
    # getContent()
    # open()
    # setContent()
    

