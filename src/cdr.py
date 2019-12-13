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
        pageConfig = []
        for page in self.doc.Pages:
            dictName = {
                "秒秒学板块":True,
                "秒秒学装饰":True,
                "秒秒学结构":True,
                "秒秒学背景":True,
                "秒秒学全局参数":True,
            }
            for curLayer in page.AllLayers:
                if dictName.get(curLayer.Name) == True:
                    dictName[curLayer.Name] = False
            
            pageConfig.append(dictName)

        for index in range(len(pageConfig)):
            for key in pageConfig[index]:
                if pageConfig[index][key] == True: 
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
        sh.Fill.UniformColor.RGBAssign(style['background-color'][0],style['background-color'][1],style['background-color'][2])
        sh.PositionX = positionX 
        sh.PositionY = positionY

