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

    def __setImage(self, determine, allLayers):
        visibleLayerName = determine.getVisibleField()
        for curLayer in allLayers:
            # 设置图片
            Input.accessImage(self.doc,  curLayer.Shapes)
            # 设置状态，处理层级可见性
            determine.setLayerVisible(curLayer, visibleLayerName)

    def __accessExtractTextData(self, pageObj, pageIndex):
        allLayers = pageObj.AllLayers
        determine = Determine()
        self.__preprocess(determine, allLayers, pageIndex)
        self.__accessInput(determine, allLayers, pageIndex)
        # 设置图片/层的可见性
        if prarm.cmdCommand == "set:text":
            self.__setImage(determine, allLayers)

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

    def get(self, pageIndex=""):
        prarm.setCommand("get:text")
        self.__accessData(pageIndex)
        return retrunData()

    def set(self, newData, pageIndex=""):
        prarm.setCommand("set:text")
        prarm.setExternalData(newData)
        self.__accessData(pageIndex)

    def drawDecorationTriangle(self):
        self.doc.Unit = 5
        ActiveLayer = self.doc.ActiveLayer
        # s1 = ActiveLayer.CreateRectangle2(0, 0, 3, 1)
        # s1.Fill.UniformColor.RGBAssign(255, 0, 0)
        # ActiveLayer.CreateCustomShape("Table", 1, 10, 5, 7, 7, 6)
        s1 = ActiveLayer.CreatePolygon(0, 100, 300, 0, 3, 1)
        s1.Fill.UniformColor.RGBAssign(255, 0, 0)
        s1.PositionX = 0
        s1.PositionY = 0
