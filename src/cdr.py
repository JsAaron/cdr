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
        setPageTotal(self.doc.Pages.Count)


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
        spath.AppendLineSegment(0, -300)
        spath.AppendLineSegment(300, 0)
        spath.Closed = True


        layer = ActivePage.CreateLayer("三角形")
        sh = layer.CreateCurve(crv)
        sh.Fill.UniformColor.RGBAssign(255, 0, 0)
        
        self.doc.ReferencePoint = 3
        sh.PositionX = 0 
        sh.PositionY = sizeheight

        # # 左上角
        # if position == 'lefttop':
    
        #     spath = crv.CreateSubPath(0, 0)
        #     spath.AppendLineSegment(0, 300)
        #     spath.AppendLineSegment(300, 0)

    
  

        # spath.Closed = True
   

