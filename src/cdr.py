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
from prarm import setExternalData, setCommand


class CDR():

    def __init__(self, path=""):
        self.app = Dispatch('CorelDraw.Application')

        if path:
            self.app.OpenDocument(path)

        self.doc = self.app.ActiveDocument

        setPageTotal(self.doc.Pages.Count)

    # 预处理
    def __preprocess(self, determine, allLayers, pageIndex):
        for curLayer in allLayers:
            determine.initField(curLayer.Name, curLayer.Shapes, pageIndex)

    # 读/取操作
    def __accessInput(self, allLayers, determine, pageIndex):
        for curLayer in allLayers:
            Input.accessShape(self.doc,  curLayer.Shapes, determine, pageIndex)

    # 获取文档所有页面、所有图层、所有图形对象
    def __accessExtractTextData(self, pageObj, pageIndex):
        allLayers = pageObj.AllLayers
        determine = Determine()
        self.__preprocess(determine, allLayers, pageIndex)
        self.__accessInput(allLayers, determine, pageIndex)

    def __accessData(self, pageIndex):
        if pageIndex:
            self.__accessExtractTextData(
                self.doc.Pages.Item(pageIndex), pageIndex)
        else:
            count = 1
            for page in self.doc.Pages:
                self.__accessExtractTextData(page, count)
                count += 1

    # 获取所有内容
    # page指定页码
    def get(self, pageIndex=""):
        setCommand("get:text")
        self.__accessData(pageIndex)
        return retrunData()

    def set(self, newData, pageIndex=""):
        setCommand("set:text")
        setExternalData(newData)
        self.__accessData(pageIndex)
