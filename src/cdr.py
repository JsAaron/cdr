import subprocess
import sys
import json
import win32file
import win32api
import win32con
import win32com.client
from win32com.client import Dispatch, constants
from determine import Determine
from input import accessShape
from result import retrunData,setPageTotal

class CDR():
    def __init__(self, path=""):
        self.app = Dispatch('CorelDraw.Application')

        if path:
            self.app.OpenDocument(path)
        self.doc = self.app.ActiveDocument

        if self.doc == None:
            self.__return("false", "文档打开失败")

        setPageTotal(self.doc.Pages.Count)

    # 定义返回
    def __return(self, status, content):
        return {
            "status ": status,
            "content": content
        }

    # 预处理
    def __preprocess(self, determine, allLayers, pageIndex):
        for curLayer in allLayers:
            determine.initField(curLayer.Name, curLayer.Shapes, pageIndex)

    # 读/取操作
    def __accessInput(self, allLayers, determine, pageIndex):
        for curLayer in allLayers:
            accessShape(self.doc,  curLayer.Shapes, determine, pageIndex)

    # 获取文档所有页面、所有图层、所有图形对象
    def __accessExtractTextData(self, pageObj, pageIndex):
        allLayers = pageObj.AllLayers
        determine = Determine()
        self.__preprocess(determine, allLayers, pageIndex)
        self.__accessInput(allLayers, determine, pageIndex)

    # 获取所有内容
    # page指定页码
    def getPageContent(self, pageIndex=""):
        if pageIndex:
            self.__accessExtractTextData(
                self.doc.Pages.Item(pageIndex), pageIndex)

        return retrunData()