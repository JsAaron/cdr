import subprocess
import sys
import json
import win32file
import win32api
import win32con
import win32com.client
from win32com.client import Dispatch, constants
from determine import Determine
import input as Input
from result import retrunData, setPageTotal
import prarm


class CDR():

    def __init__(self, path=""):
        self.app = Dispatch('CorelDraw.Application')
        if path:
            self.app.OpenDocument(path)
        self.doc = self.app.ActiveDocument
        self.__initDefalutLayer()
        setPageTotal(self.doc.Pages.Count)

    # 初始默认图层
    def __initDefalutLayer(self):
        pagesConfig = []
        for page in self.doc.Pages:
            dictName = {
                "秒秒学板块": True,
                "秒秒学装饰": True,
                "秒秒学结构": True,
                "秒秒学背景": True,
                "秒秒学全局参数": True,
            }

            for curLayer in page.AllLayers:
                if dictName.get(curLayer.Name) == True:
                    dictName[curLayer.Name] = False

            pagesConfig.append(dictName)

        for index in range(len(pagesConfig)):
            pageCfg = pagesConfig[index]
            for key in pageCfg:
                if pageCfg[key] == True:
                   has = False
                   cpage = self.doc.Pages.Item(index+1)
                    # 多页面第二次去重
                   for curLayer in cpage.AllLayers:
                        if curLayer.Name == key:
                            has = True

                   if has == False:
                    self.doc.Pages.Item(index+1).CreateLayer(key)

    def __preprocess(self, determine, allLayers, pageIndex):
        for curLayer in allLayers:
            determine.initField(curLayer.Name, curLayer.Shapes, pageIndex)


    def __accessInput(self,  determine, allLayers, pageIndex):
        for curLayer in allLayers:
            Input.accessShape(self.doc,  curLayer.Shapes, determine, pageIndex)


    def __setImage(self, determine, allLayers, pageIndex):
        visibleLayerName = determine.getVisibleField()
        for curLayer in allLayers:
            # 设置图片
            Input.accessImage(self.doc, curLayer.Shapes, pageIndex)
            # 设置状态，处理层级可见性
            determine.setLayerVisible(curLayer, visibleLayerName)


    def __accessExtractTextData(self, pageObj, pageIndex):
        allLayers = pageObj.AllLayers
        determine = Determine()
        self.__preprocess(determine, allLayers, pageIndex)
        self.__accessInput(determine, allLayers, pageIndex)
        # 设置图片/层的可见性
        if prarm.cmdCommand == "set:text":
            self.__setImage(determine, allLayers, pageIndex)


    def __accessData(self, pageIndex):
        if pageIndex:
            self.__accessExtractTextData(
                self.doc.Pages.Item(pageIndex), pageIndex)
        else:
            count = 1
            for page in self.doc.Pages:
                self.__accessExtractTextData(page, count)
                count += 1


    # 根据名称找到图层
    def __getAssignLayer(self,name):
       for curLayer in self.doc.ActivePage.AllLayers:
            if curLayer.Name == name:
                return curLayer


    # =================================== 对外 ===================================

    # 切换页面
    def togglePage(self,pageIndex=1):
        if pageIndex == 0:
           print("pageIndex不能为0")
           return

        if pageIndex > self.doc.Pages.Count:
           print("设置页码数大于总页数")
           return
        self.doc.Pages.Item(pageIndex).Activate()
        return self.get(pageIndex)


    # 获取所有数据段
    def get(self, pageIndex=""):
        prarm.setCommand("get:text")
        self.__accessData(pageIndex)
        return retrunData()


    # 更新数据段
    def set(self, newData, pageIndex=""):
        prarm.setCommand("set:text")
        prarm.setExternalData(newData)
        self.__accessData(pageIndex)


    def groupDecorationTriangle(self):
        sh1 = self.drawDecorationTriangle("test",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   
        sh2 = self.drawDecorationTriangle("test",{"background-color":[255, 0, 0]},{"bottom":300,"right":600},'righttop') 
        sr = self.app.ActiveSelection.Shapes
        for key in sr:
            key.Layer = self.doc.ActiveLayer
            key.Group(sh1)


    # 创建边界三角形
    def drawDecorationTriangle(self, name, style, points, position):
        self.doc.Unit = 5

        ActivePage = self.doc.ActivePage
        sizeheight = ActivePage.sizeheight
        sizewidth = ActivePage.sizewidth
     
        crv = self.app.CreateCurve(self.doc)
        spath = crv.CreateSubPath(0, 0)

        x = 0
        y = 0
        positionX = 0
        positionY = 0 

       # 左上角
        if position == 'lefttop':
            self.doc.ReferencePoint = 3
            x = -points['bottom']
            y = points['left']
            positionX = 0 
            positionY = sizeheight
        
        if position == 'righttop':
            self.doc.ReferencePoint = 1
            x = -points['bottom']
            y = -points['right']
            positionX = sizewidth 
            positionY = sizeheight

        if position == 'leftbottom':
            self.doc.ReferencePoint = 5
            x = points['top']
            y = points['left']

        if position == 'rightbottom':
            self.doc.ReferencePoint = 7
            x = points['top']
            y = -points['right']
            positionX = sizewidth 

        spath.AppendLineSegment(0, x)
        spath.AppendLineSegment(y, 0)
        spath.Closed = True

        layer = self.__getAssignLayer("秒秒学装饰")
        sh = layer.CreateCurve(crv)
        sh.Name = name
        sh.Fill.UniformColor.RGBAssign(style['background-color'][0],style['background-color'][1],style['background-color'][2])
        sh.PositionX = positionX 
        sh.PositionY = positionY
        return sh


    #合并多个形状分组
    # layer 指定层
    # name 新的分组名字
    # [s1,s2,s3...] 需要合并的对象明数组
    def groupShape(self,layer,name,shapeNames):
        parents = layer.FindShape(name)
        if parents != None:
           return parents

        groupIndex = []
        for index in range(len(layer.Shapes)):
            itemIndex = index+1
            item = layer.Shapes.Item(itemIndex)
            if item.Name in shapeNames:
                groupIndex.append(itemIndex)
                
        rs = layer.Shapes.Range(groupIndex)
        g = rs.Group()
        g.Name = name
        return g

    # def addGroupShape(self,original,target):





    # 分栏文本
    def insertColumnText(self):
        layer = self.__getAssignLayer("秒秒学装饰")
        s1 =  layer.FindShape("test1")
        s2 =  layer.FindShape("test2")
        s3 =  layer.FindShape("test3")
        s4 =  layer.FindShape("test4")

        #  Shape.Group
        if s1 == None:
            s1 = self.drawDecorationTriangle("test1",{"background-color":[255, 0, 0]},{"bottom":300,"left":600},'lefttop')   
        
        if s2 == None:
            s2 = self.drawDecorationTriangle("test2",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')   
     
        if s3 == None:
            s3 = self.drawDecorationTriangle("test3",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')   

        if s4 == None:
            s4 = self.drawDecorationTriangle("test4",{"background-color":[255, 0, 0]},{"top":300,"right":600},'rightbottom')   


        groups = self.groupShape(layer,"测试群1",['test2','test3'])

        # groups.Ungroup()
        # s4.AddToPowerClip(groups)
        # s1.AddToSelection (groups)
        # self.addShapeToGroup(groups,'test4')
 
