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


# 移动对象
def moveToMiddle():
    cdrObj = CDR()
    # cdrObj.moveToLandscapeMiddle("测试1")
    # cdrObj.moveToLeft(obj)
    # cdrObj.moveToRight('测试1')
    # cdrObj.moveToTop('测试1')
    # cdrObj.moveToBottom('测试1')
    # cdrObj.moveToVerticalMiddle('测试1')
    cdrObj.moveToCenter('测试1')


# 字体尺寸修改
def increaseFontSize():
    cdrObj = CDR()
    cdrObj.loadPalette('C:\\Users\\Administrator\\Desktop\\123\\cw.xml')
    # cdrObj.addFontSize('测试1',24)
    #  cdrObj.reduceFontSize('测试1',15)
    # cdrObj.setColor('测试1',[255,0,0])

def testColor():
    cdrObj = CDR()
    # cdrObj.setColor('测试1',[255,0,0])
    cdrObj.createPage(5)


# 组合测试
def combineTest():
    cdrObj = CDR()
    obj = cdrObj.insertParaText([0, 0, 120, 500],'测试1','我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf1111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf111111')
    cdrObj.setColor(obj,[255,0,0])
    cdrObj.addFontSize(obj,24)
    cdrObj.moveToCenter(obj)
    cdrObj.modifyParaText(obj,'dfs123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111我是内容123123大1sdfasfsfsdf2312dfsadfsfsdf11111',[5, 50, 100, 50],'','')
    cdrObj.setFontSize(obj,10)


    #创建调色版，并增加颜色对象
def paletteTest1():
    cdrObj = CDR()
    paletteObj = cdrObj.accessPalette('my') 
    cdrObj.setPletteEnabled(paletteObj) #启用
    # 创建一个颜色对象,使用指定格式
    color = cdrObj.createColorObj([110,128,255],'unique_key','RGB')  
    # 把颜色增加到调色板上
    cdrObj.addPletteColor(paletteObj,color)


#测试调色板, 替换颜色
def paletteTest2():
    cdrObj = CDR()
    paletteObj = cdrObj.accessPalette('my') 
    newColor = cdrObj.createColorObj([0,255,255],'unique_key','RGB')  
    cdrObj.replacePletteColorByName(paletteObj,newColor)


# 测试调色板,获取调色板中的指定颜色
def paletteTest3():
    cdrObj = CDR()
    print(cdrObj.app.Printers)

    # cdrObj.test()
    # paletteObj = cdrObj.accessPalette('test-paletter')
    # cdrObj.setPletteDefault(paletteObj)

    # 获取颜色对象 
    # colorObj = cdrObj.getPaletteColor(paletteObj,'unique_key')
    # 获取值
    # cmykValue = cdrObj.getColorValue(colorObj,'CMYK')
    # rgbValue = cdrObj.getColorValue(colorObj,'RGB')
    # hsbValue = cdrObj.getColorValue(colorObj,'HSB')
    # hlsValue = cdrObj.getColorValue(colorObj,'HLS')
    # cmkValue = cdrObj.getColorValue(colorObj,'CMY')
    # print(colorObj)


# 导入文件，并替换对象
def replacePart():
    cdrObj = CDR()
    # 加载路径下的cdr文件，中的mytest对象
    # 替换到指定的对象
    # cdrObj.replacePart(['C:\\Users\\Administrator\\Desktop\\111\\2.cdr','mytest'],cdrObj.app.ActiveShape)


# 保存文件
def testSaveCDR():
    cdrObj = CDR()

    # 导出指定页面
    cdrObj.exportBitmap(1,1136,700,'C:\\Users\\Administrator\\Desktop\\111\\test.jpg')
    
    # 导出所有页面
    # 只要目录
    # cdrObj.exportAllBitmap(1136,700,'C:\\Users\\Administrator\\Desktop\\111\\')
    # cdrObj.saveCDR('C:\\Users\\Administrator\\Desktop\\111\\51.cdr')


if __name__ == '__main__':
    # testSaveCDR()
    # testSaveCDR()
    # importText()
    # paletteTest1()
    # paletteTest2()
    testSaveCDR()

    
    

