import subprocess
import sys
import win32com.client
from win32com.client import Dispatch, constants, GetActiveObject
from determine import Determine
import input as Input
from result import retrunData, setPageTotal
import prarm
import urllib.parse
import os
import time


DEFAULTLINEHEIGHT = 5.5  # mm


class CDR():

    def __init__(self, path=""):
        try:
            self.app = GetActiveObject('CorelDraw.Application')
        except:
            self.app = Dispatch('CorelDraw.Application')
        if path:
            self.app.OpenDocument(path)
        self.doc = self.app.ActiveDocument
        self.doc.Unit = 3               # mm unit
        self.pagewidth = self.doc.ActivePage.SizeWidth
        self.pageheight = self.doc.ActivePage.SizeHeight
        self.palette = self.doc.Palette
        self.__initDefalutLayer()
        self.togglePage(1)
        setPageTotal(self.doc.Pages.Count)
        # adjust original point to top for double-page
        if self.doc.ActivePage.TOPY > 0:
            self.doc.DrawingOriginY = self.doc.ActivePage.TOPY / 2

    # 初始默认图层
    def __initDefalutLayer(self):

        # 母版
        hasMasterLayer = False
        for masterLayer in self.doc.MasterPage.AllLayers:
            if masterLayer.Name == '秒秒学全局参数':
                hasMasterLayer = True
        if hasMasterLayer == False:
            self.doc.MasterPage.createlayer('秒秒学全局参数')

        # 页面
        pagesConfig = []
        for page in self.doc.Pages:
            dictName = {
                "秒秒学背景": True,
                "秒秒学板块": True,
                "秒秒学结构": True,
                "秒秒学装饰": True
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

     # 探测图片是否已经创建
    # 默认探测5次

    def __detectionImage(self, layer, imageName, count=10):
        obj = layer.FindShape(imageName)
        # 探测结束
        if count == 0:
            return obj
        if obj == None:
            time.sleep(0.1)
            count = count-1
            return self.__detectionImage(layer, imageName, count)
        else:
            return obj

    # 移动形状到缓存
    def __moveShapeToCache(self, layerObj, shapeObj):
        delGroupObj = self.createDeleteCache(layerObj)
        firstObj = delGroupObj.Shapes.Item(1)
        shapeObj.OrderFrontOf(firstObj)
        delGroupObj.Delete()

    # =================================== 基础方法 ===================================

    # 公开创建标准目录层接口
    def createStdFolder(self):
        return self.__initDefalutLayer()


    # 判断变量类型
    def getType(self, variate):
        type = None
        if isinstance(variate, int):
            type = "int"
        elif isinstance(variate, str):
            type = "str"
        elif isinstance(variate, float):
            type = "float"
        elif isinstance(variate, list):
            type = "list"
        elif isinstance(variate, tuple):
            type = "tuple"
        elif isinstance(variate, dict):
            type = "dict"
        return type


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


    # name 根据名称找到图层
    # page 指定页面搜索layer
    def findLayerByName(self, name=''):
        if name:
            for curLayer in self.doc.ActivePage.AllLayers:
                if curLayer.Name == name:
                    return curLayer
        return self.doc.ActiveLayer


    # 通过ID过去形状对象
    def findShapeById(self, parentObj, shapeObj):
        groupObj = None
        for shape in parentObj.Shapes:
            if shape.StaticID == shapeObj.StaticID:
                groupObj = shape
                break
        return groupObj
    
    # 通过id找到相应的对象
    # 增加对组的处理 xiaowy 2019/12/25
    def findShapeById2(self, parentobj ,objId):
        groupObj = None
        for shape in parentobj.Shapes:
            # if shape.Name == parentobj:
            if shape.StaticID == objId:
                # groupObj = shape
                # break
                return shape
            elif not shape.IsSimpleShape: # for group
                groupObj = self.findShapeById2(shape, objId)
        return groupObj



    # 找到当前组内的形状
    # 增加对组的处理 xiaowy 2019/12/25
    def findShapeByName(self, parentobj ,name):
        groupObj = None
        for shape in parentobj.Shapes:
            # if shape.Name == parentobj:
            if shape.Name == name:
                # groupObj = shape
                # break
                return shape
            elif not shape.IsSimpleShape: # for group
                groupObj = self.findShapeByName(shape, name)
        return groupObj


    # 找到当前组内的形状合集
    def findShapeByNames(self, name, parentobj=None):
        groupShapes = []
        for shape in parentobj.Shapes:
            if shape.Name == name:
                groupShapes.append(shape)
        return groupShapes


    # 转化成指定对象
    # name id obj
    def transformObjs(self, shapeObj):
        if self.getType(shapeObj) == 'str':
            layer = self.findLayerByName()
            shapeObj = layer.FindShape(shapeObj)
        return shapeObj


    # 切换页面
    def togglePage(self, pageIndex=1):
        if pageIndex == 0:
           print("pageIndex不能为0")
           return
        if pageIndex > self.doc.Pages.Count:
           print("设置页码数大于总页数")
           return
        if self.app.ActivePage.Index == pageIndex:
           return
        page = self.doc.Pages.Item(pageIndex)
        page.Activate()
        return page


    # select predefined layer
    def selectLayer(self, layername):
        # 必须设置活动的layer，这样调用vb.exe才会在这个layer的内部
        layer = self.findLayerByName(layername)
        if layer != None:
            layer.Activate()
        return layer


    # 添加图片
    # imagePath："C:\\Users\\Administrator\\Desktop\\111\\1.png"
    def addImage(self, layer, imagePath):
        # 路径转码
        data = "{'path':'" + urllib.parse.quote(imagePath) + "'}"
        parent = os.path.dirname(os.path.realpath(__file__))
        vbPath = parent + '\\vb\\ConsoleApp.exe'
        # 参数只有一个路径
        # data = "{'path':'C%3A%5CUsers%5CAdministrator%5CDesktop%5C111%5C1.png'}"
        cmdStr = [vbPath, 'add:image', data]
        subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
        return self.__detectionImage(layer, os.path.basename(imagePath))


    # inesert placeholder, all placeholder is the same name
    def insertPlaceholder(self, layerObj):
        placeholderObj = layerObj.CreateLineSegment(10, 10, 11, 11)
        placeholderObj.Name = "placeholder"
        placeholderObj.Outline.Type = 0
        return placeholderObj


    # move shapes into existing groupobj
    def moveShapeToGroup(self, groupobj, shapeObjs):
        firstmember = groupobj.Shapes.Item(1)
        allIds = [k.StaticID for k in groupobj.Shapes]
        for shape in shapeObjs:
            if shape.StaticID in allIds:  # shape already in group
                continue
            else:
                shape.OrderFrontOf(firstmember)
        return groupobj


    # find whether a shape is in a group
    def findShapeInGroup(self, groupobj, shapename):
        for shape in groupobj.Shapes:
            if shape.Name == shapename:
                return shape
        return None


    # 合并多个形状分组
    # layer 指定层
    # name 新的分组名字
    # [s1,s2,s3...] 需要合并的对象名称数组
    # shapeObjs must in layerObj, but maynot in parentobj
    def groupShapeObjs(self, layerObj, groupName, parentobj=None, shapeObjs=[]):
        # first try to find the group, shape object has not Findshape method
        if parentobj == None:
            parentobj = layerObj

        groupObj = self.findShapeByName(parentobj, groupName)

        # new group must has at least two shapeObjs
        if groupObj == None:

            # if shapeObjs dont have enough objects, add placeholder
            if len(shapeObjs) < 1:
                shapeObjs.append(self.insertPlaceholder(layerObj))
            if len(shapeObjs) < 2:
                # the placeholder for one object should not change the group'size, so move it to the same position with
                # # object
                theplaceholder = self.insertPlaceholder(layerObj)
                orishapeobj = shapeObjs[0]
                theplaceholder.PositionX = orishapeobj.PositionX
                theplaceholder.PositionY = orishapeobj.PositionY
                shapeObjs.append(theplaceholder)

            if len(shapeObjs) < 2:  # should not happen because we add placeholder, anyway keep it for safe
                return None
            else:  # create new group
                # if it does has parentgroup, move to parent first
                if parentobj != layerObj:
                    self.moveShapeToGroup(parentobj, shapeObjs)
                # then make new group
                groupIndex = []
                shapeIDs = [k.StaticID for k in shapeObjs]
                for index in range(len(parentobj.Shapes)):
                    itemIndex = index + 1
                    item = parentobj.Shapes.Item(itemIndex)
                    if item.StaticID in shapeIDs:
                        groupIndex.append(itemIndex)

                rs = parentobj.Shapes.Range(groupIndex)
                newGroup = rs.Group()
                if newGroup != None:
                    newGroup.Name = groupName
                return newGroup
        else:
            # Already has group
            return self.moveShapeToGroup(groupObj, shapeObjs)


    # 创建临时删除区域
    # 删除不能用Delete, 直接delete会把整体列表都移除
    # 需要创建一个临时的对象，把shapre移动过去，最后删除这个临时对象
    def createDeleteCache(self, layerObj):
        shapeObjs = []
        groupIndex = []
        shapeObjs.append(self.insertPlaceholder(layerObj))
        shapeObjs.append(self.insertPlaceholder(layerObj))
        for index in range(len(layerObj.Shapes)):
            itemIndex = index + 1
            item = layerObj.Shapes.Item(itemIndex)
            if item.Name == 'placeholder':
                groupIndex.append(itemIndex)
        rs = layerObj.Shapes.Range(groupIndex)
        delGroupObj = rs.Group()
        delGroupObj.Name = "临时删除组"
        return delGroupObj


    # 增加形状对象到组对象
    def addShapeToGroup(self, groupObj, shapeObj):
        # 本身不是组结构
        if groupObj.Shapes.Count == 0:
            return
        firstmember = groupObj.Shapes.Item(1)
        shapeObj.OrderFrontOf(firstmember)
        placeholderArr = self.findShapeByNames('placeholder', groupObj)
        # 如果当前组下还存在预创建对象
        if len(placeholderArr):
            delGroupObj = self.createDeleteCache(groupObj.Layer)
            firstObj = delGroupObj.Shapes.Item(1)
            for index in range(len(placeholderArr)):
                item = placeholderArr[index]
                if item.Name == 'placeholder':
                    item.OrderFrontOf(firstObj)
            delGroupObj.Delete()
        return groupObj


    # 从组中移除指定的对象
    # 保持组的持久
    def removGroupShapeObjs(self, groupObj, shapeObj):
        hasObj = self.findShapeById(groupObj, shapeObj)
        if hasObj == None:
            return hasObj
        layerObj = groupObj.Layer
        # 必须保持结构
        if groupObj.Shapes.Count == 1:
            # 新加入占位
            self.insertPlaceholder(layerObj).OrderFrontOf(shapeObj)
            self.insertPlaceholder(layerObj).OrderFrontOf(shapeObj)
        self.__moveShapeToCache(layerObj, shapeObj)
        return groupObj


    # 删除组对象，如果组为空,不保持组的存在
    def deleteGroupShapeObjs(self, groupObj, shapeObj):
        hasObj = self.findShapeById(groupObj, shapeObj)
        if hasObj == None:
            return hasObj
        layerObj = groupObj.Layer
        self.__moveShapeToCache(layerObj, shapeObj)
        if groupObj.Shapes.Count == 0:
            groupObj.Ungroup()
        return groupObj


    # 复制对象
    # obj是一个对象，也可以是一个组
    # newname是复制后对象的名字
    def cloneShape(self, obj, newname, OffsetX=0, OffsetY=0):
        newObj = obj.Duplicate(OffsetX, OffsetY)
        newObj.Name = newname
        return newObj


    # copy a block from master Layer to current layer
    def cloneFromMaster(self, layer, groupname):
        # select the group in master layer
        masterpage = self.doc.MasterPage
        masterlayer = None
        for mlayer in masterpage.AllLayers:
            if mlayer.Name == '板块':
                masterlayer = mlayer
                break
        if masterlayer == None:
            return None
        thegroup = masterlayer.FindShape(groupname, 7)      # cdrGroupShape
        if thegroup == None:
            return None
        # copy it
        thegroup.CopyToLayer(layer)
        copiedgroup = layer.FindShape(groupname)
        return copiedgroup

    # find all the block templates of a type

    # ========================== 创建/修改 ==========================

    def groupDecorationTriangle(self):
        sh1 = self.drawDecorationTriangle(
            "test", {"background-color": [255, 0, 0]}, {"bottom": 300, "left": 600}, 'lefttop')
        sh2 = self.drawDecorationTriangle(
            "test", {"background-color": [255, 0, 0]}, {"bottom": 300, "right": 600}, 'righttop')
        sr = self.app.ActiveSelection.Shapes
        for key in sr:
            key.Layer = self.doc.ActiveLayer
            key.Group(sh1)


    # 创建边界三角形
    def drawDecorationTriangle(self, name, style, points, position):

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

        # layer = self.getLayer("秒秒学装饰")
        layer = self.doc.ActivePage.ActiveLayer
        sh = layer.CreateCurve(crv)
        # sh.Name = '三角形' + position
        # sh.Name = name
        # print(sh.Name)
        sh.Fill.UniformColor.RGBAssign(
            style['background-color'][0], style['background-color'][1], style['background-color'][2])
        sh.PositionX = positionX
        sh.PositionY = positionY
        sh.Name = name
        return sh


    # adjust the height of the paragraph text, until all the content has been displayed
    def adjustParaTextHeight(self, textobj):
        if textobj.Type == 6:           # cdrTextShape
            if textobj.Text.Type == 1:  # cdrParagraphText
                if textobj.Text.Overflow:
                    amount = DEFAULTLINEHEIGHT     # 15.6pt, this is a fixed line height grid
                    maxloop = 100    # double height of all the A4 page. Means we deal most double A4 page
                    initloop = 0
                    for i in range(maxloop):
                        if textobj.Text.Overflow:
                            textobj.SizeHeight += amount
                        else:
                            break
        pass


    # convert coordinates
    # in common sense, the coordinate should start at left top and y axis is down, so is is the think of people
    # which write from left to right, from top to bottom
    # in Coreldraw, it is different:
    # single-page: left -> right is the same, but original point is left top, y axis is up
    # double-page: left is different for left-page or right-page, and original point is bottom center, y axis is up
    def convertCood(self, oribound):
        bound = oribound.copy()
        lx = self.doc.ActivePage.LeftX
        if lx != 0:
            bound[0] += lx
            bound[2] += lx
        return bound


    # revert clac
    def revertCood(self, xvalue):
        lx = self.doc.ActivePage.LeftX
        return xvalue - lx

    # insert paragraph text
    # bound: text bound, height will be auto calc, maybe not in bound
    # style: text style of the paragraph
    # content: text string
    # the height of paragraph text should only one line, because we calc overflow, only enlarge, not shrink
    def insertParaText(self, oribound, name='正文', content='', style='', paletteidx=2, shape=None, columns = 0):
        # if the text already exist, just adjust it's bound
        theobj = shape
        newHeight = 0
        if theobj != None:
            # adjust it's bound
            story = theobj.Text.Story.Text
            # for unknown, can't simple replace content, or there will be font error
            # if story != content:
            #    theobj.Text.Replace(story, content, True)
            #    # now the height changes, we should totaly reassign the height
            #    newHeight = DEFAULTLINEHEIGHT
            self.moveObj(theobj, oribound)
            # only width change exceed half characters, should we really change text width
            # 2.5 is in mm, and half characters
            if abs(theobj.SizeWidth - (oribound[2] - oribound[0])) > DEFAULTLINEHEIGHT/2 or newHeight > 0:
                theobj.SizeWidth = oribound[2] - oribound[0]
                theobj.SizeHeight = DEFAULTLINEHEIGHT
                self.adjustParaTextHeight(theobj)
            # xiaowy
            if columns > 0:
                self.setTextColumns(theobj, columns)
            return theobj
        bound = self.convertCood(oribound)
        theobj = self.doc.ActiveLayer.CreateParagraphText(
            bound[0], -1 * bound[1], bound[2], -1 * bound[1] - 1, content, 0, -1)
        if style:
            theobj.ApplyStyle(style)
        if self.palette.ColorCount >= paletteidx:
            theobj.Text.Story.Fill.UniformColor = self.palette.Colors()[paletteidx]
        theobj.Name = name
        theobj.SizeHeight = DEFAULTLINEHEIGHT
        self.adjustParaTextHeight(theobj)
        # xiaowy
        if columns > 0:
                self.setTextColumns(theobj, columns)
        return theobj

    def setTextColumns(self, textObj, columns):
        '''
        设定文本框的分栏
        key-words:
            textObj:text 对象
            columns:分栏的个数
        '''
        story = textObj.Text.Story
        fontSize = story.Size
        columnSpace = 1.5 * fontSize
        columnWidth = (textObj.SizeWidth - (columns - 1) * columnSpace) // columns
        frame = textObj.Text.Frame
        frame.SetColumns(columns, True, [columnWidth, columnSpace])
        return textObj

    # 修改段落文本
    # shapeObj 文本对象
    # content 文本内容
    # oribound 坐标系
    # style
    # paletteidx 调色表索引
    # name 节点名字
    def modifyParaText(self, shapeObj, content='', oribound=[], style='', paletteidx='', name=''):
        shapeObj = self.transformObjs(shapeObj)
        if shapeObj == None or content == '' or shapeObj.Text.Story.Text == content:
            return
        shapeObj.Text.Story.Text = ''
        shapeObj.Text.Story.Text = content
        if name:
            shapeObj.Name = name
        if style:
            shapeObj.ApplyStyle(style)
        if paletteidx and self.palette.ColorCount >= int(paletteidx):
            shapeObj.Text.Story.Fill.UniformColor = self.palette.Colors()[paletteidx]
        if len(oribound):
            self.moveObj(shapeObj, oribound)
            self.adjustParaTextHeight(shapeObj)


    # insert paragraph text
    # bound: text bound, height will be auto calc, maybe not in bound
    # style: text style of the paragraph
    # content: text string
    def insertPointText(self, oribound, name='主标题', content='', style='正文', paletteidx=2, shape=None):
        theobj = shape
        if theobj != None:
            # adjust it's bound
            story = theobj.Text.Story.Text
            # for unknown, can't simple replace content
            # if story != content:
            #    theobj.Text.Replace(story, content, True)
            self.moveObj(theobj, oribound)
            return theobj
        bound = self.convertCood(oribound)
        theobj = self.doc.ActiveLayer.CreateArtisticText(
            bound[0], -1 * bound[3], content, 0, -1)
        theobj.ApplyStyle(style)
        theobj.Text.Story.Fill.UniformColor = self.palette.Colors()[paletteidx]
        theobj.Name = name
        theobj.PositionY = -1 * bound[1]
        return theobj


    # insert background rect
    def insertRectangle(self, oribound, name='正文背景', round=0, paletteidx=0, noborder=True, shape=None):
        theobj = shape
        if theobj != None:
            # adjust it's bound
            self.moveObj(theobj, oribound)
            self.sizeObj(theobj, oribound)
            return theobj
        bound = self.convertCood(oribound)
        theobj = self.doc.ActiveLayer.CreateRectangle(bound[0], -1 * bound[1], bound[2], -1 * bound[3],
                    round, round, round, round)
        theobj.Fill.UniformColor = self.palette.Colors()[paletteidx]
        if noborder:
            theobj.ApplyStyle('无轮廓')
        theobj.Name = name
        return theobj


    # insert powerclip from rectangle
    def insertPowerclip(self, oribound, name='图片', round=0, style='图文框', shapeImg=None, shape='rect'):
        '''
        key-words:
            shapeImg:图片形状对象,默认None
            shape:形状,默认为矩形(rect),可为'circle'或'triangle'
        '''
        theobj = shapeImg
        if theobj != None:
            # adjust it's bound
            self.moveObj(theobj, oribound)
            self.sizeObj(theobj, oribound)
            return theobj
        bound = self.convertCood(oribound)
        if shape == 'rect':
            theobj = self.doc.ActiveLayer.CreateRectangle(bound[0], -1 * bound[1], bound[2], -1 * bound[3],
                        round, round, round, round)
            theobj.ApplyStyle(style)
            theobj.Name = name
            rect2 = self.doc.ActiveLayer.CreateRectangle(bound[0], bound[1], bound[2], bound[3],
                        round, round, round, round)
            # rect2.Name = 'test' # for debug
            rect2.AddToPowerClip(theobj, -1)
        
            return theobj
        elif shape == 'circle':
            theObj = self.doc.ActiveLayer.CreateEllipse(bound[0], -1 * bound[1], bound[2], -1 * bound[3])
            theObj.Name = name
            theObj.ApplyStyle(style)
            circle = self.doc.ActiveLayer.CreateEllipse(bound[0], bound[1], bound[2], bound[3])
            circle.AddToPowerClip(theObj, -1)
            return theObj
        else:
            raise NotImplementedError('Waitting for next viersion...')


    # insert line
    def insertLine(self, oribound, name='分隔线', style='粗分隔线', type='horizontal', shape=None):
        theobj = shape
        if theobj != None:
            # adjust it's bound
            self.moveObj(theobj, oribound)
            self.sizeObj(theobj, oribound, withheight=False)
            return theobj
        bound = self.convertCood(oribound)
        if type == 'horizontal':
            theobj = self.doc.ActiveLayer.CreateLineSegment(
                bound[0], -1*bound[1], bound[2], -1*bound[1])
        elif type == 'vertical':
            theobj = self.doc.ActiveLayer.CreateLineSegment(
                bound[0], -1*bound[1], bound[0], -1*bound[3])
        else:
            theobj = self.doc.ActiveLayer.CreateLineSegment(
                bound[0], -1*bound[1], bound[2], -1*bound[3])
        theobj.ApplyStyle(style)
        theobj.Name = name
        return theobj


    # set random outline for an object, mostly for block frame
    def setRandomOutline(self, shape):
        shape.Outline.Color.HLSAssign(random.randint(0, 360), 100, 100)
        pass
    

    # insert image
    def dealBackgroundImage(self, name):
        theobj = self.doc.ActiveLayer.FindShape(name)
        if theobj == None:
            return
        theobj.OrderToBack()
        theobj.PositionX = self.doc.ActivePage.LeftX
        theobj.PositionY = 0
        theobj.SizeWidth = self.doc.ActivePage.SizeWidth
        theobj.SizeHeight = self.doc.ActivePage.SizeHeight
        return theobj


    # change the size of obj
    def sizeObj(self, obj, oribound, withheight=True):
        bound = self.convertCood(oribound)
        if obj.SizeWidth != bound[2] - bound[0]:
            obj.SizeWidth = bound[2] - bound[0]
        if obj.SizeHeight != bound[3] - bound[1] and withheight:
            obj.SizeHeight = bound[3] - bound[1]


    # make obj suit for new bound
    def newbound(self, obj, newbound):
        self.moveObj(obj, newbound)
        self.sizeObj(obj, newbound)


    # align object in block frame bound
    def alignObject(self, oribound, obj, halign='center', valign='top', convert = True):
        bound = oribound.copy()
        if convert:
            bound = self.convertCood(oribound)
        objwidth = obj.SizeWidth
        objheight = obj.SizeHeight

        boundwidth = bound[2] - bound[0]
        boundheight = bound[3] - bound[1]

        left = bound[0]
        if halign == 'center':
            left = bound[0] + abs(boundwidth - objwidth) / 2
            pass
        elif halign == 'right':
            left = bound[0] + abs(boundwidth - objwidth)
            pass
        else:  # left
            pass

        if obj.PositionX != left:
            obj.PositionX = left

        top = - bound[1]
        if halign == 'middle':
            top = -(bound[1] + abs(boundheight - objheight) / 2)
            pass
        elif halign == 'bottom':
            top = -(bound[1] + abs(boundheight - objheight))
            pass
        else:  # top
            pass

        if valign != None:
            if obj.PositionY != top:
                obj.PositionY = top
        pass

    # align all the objects in a group
    def alignObjectInGroup(self, groupobj, halign='center', valign='top'):
        groupbound = [groupobj.LeftX, -groupobj.TOPY, groupobj.RightX, -groupobj.bottomY]
        for shape in groupobj.Shapes:
            self.alignObject(groupbound, shape, halign, None, convert = False)
        pass


    
    # ========================== 移动 ==========================

    # move object
    def moveObj(self, obj, oribound):
        bound = self.convertCood(oribound)
        if obj.PositionX != bound[0]:
            obj.PositionX = bound[0]
        if obj.PositionY != -1 * bound[1]:
            obj.PositionY = -1 * bound[1]


    # 移动X轴
    def moveLandscapeObj(self, obj, amount):
        obj.PositionX = amount


    # 移动Y轴
    def moveVerticalObj(self, obj, amount):
        obj.PositionY = -1 * amount


    # movedown object
    def movedownObj(self, obj, amount):
        if amount != 0:
            obj.PositionY -= amount


    # 移动左边
    def moveToLeft(self, name):
        sh = self.transformObjs(name)
        self.moveLandscapeObj(sh, 0)


   # 移动到中间
    def moveToLandscapeMiddle(self, name):
        sh = self.transformObjs(name)
        self.moveLandscapeObj(sh, (self.pagewidth - sh.SizeWidth)/2)


    # 移动右边
    def moveToRight(self, name):
        sh = self.transformObjs(name)
        self.moveLandscapeObj(sh, self.pagewidth - sh.SizeWidth)


    # 移动到顶部
    def moveToTop(self, name):
        sh = self.transformObjs(name)
        self.moveVerticalObj(sh, 0)


   # 移动到垂直中间
    def moveToVerticalMiddle(self, name):
        sh = self.transformObjs(name)
        self.moveVerticalObj(sh, (self.pageheight - sh.SizeHeight)/2)


    # 移动到垂直底部
    def moveToBottom(self, name):
        sh = self.transformObjs(name)
        self.moveVerticalObj(sh, self.pageheight - sh.SizeHeight)


    # 移动到正中间
    def moveToCenter(self, name):
        sh = self.transformObjs(name)
        self.moveObj(sh, [(self.pagewidth - sh.SizeWidth) /
                     2, (self.pageheight - sh.SizeHeight)/2])



    # ========================== 尺寸 ==========================


    #设置字体尺寸
    def setFontSize(self,shapeObj,value):
        shapeObj.Text.Story.Size = value
        shapeObj.SizeHeight = DEFAULTLINEHEIGHT
        if shapeObj.Text.Overflow:
            self.adjustParaTextHeight(shapeObj)


    # 增大字体
    # shapeObj 对象
    # value 设置的值
    # baseValue 如果存在基础值
    def addFontSize(self, shapeObj, value, baseValue = ''):
        shapeObj = self.transformObjs(shapeObj)
        oldSize = shapeObj.Text.Story.Size
        if baseValue:
            self.setFontSize(shapeObj,oldSize - baseValue)
        else:
            if value > oldSize:
                 self.setFontSize(shapeObj,value)


    # 减小字体
    # shapeObj 对象
    # value 设置的值
    # baseValue 如果存在基础值
    def reduceFontSize(self,shapeObj,value,baseValue = ''):
        shapeObj = self.transformObjs(shapeObj)
        oldSize = shapeObj.Text.Story.Size
        if baseValue:
            self.setFontSize(shapeObj, oldSize + baseValue)
        else:
            if value < oldSize:
                self.setFontSize(shapeObj,value)
 

    # 字体自动递增
    # shapeObj 对象
    # baseValue 基础值
    def increaseFontSize(self,shapeObj,baseValue):    
        self.addFontSize(shapeObj,'',baseValue) 

    
    # 自动递减
    # shapeObj 对象
    # baseValue 基础值
    def decreaseFontSize(self,shapeObj,baseValue):
        self.reduceFontSize(shapeObj,'',baseValue) 



    # ========================== 调色板配色 ==========================

    # 设置颜色
    def setColor(self,shapObj,rgb=[]):
        shapeObj = self.transformObjs(shapObj)
        shapeObj.Fill.UniformColor.RGBAssign(rgb[0], rgb[1], rgb[2])


    # 获取颜色值
    # colorObj 颜色对象
    # mode 返回的颜色模式
    def getColorValue(self,colorObj,mode="RGB"):
        if mode == 'RGB':
            if colorObj.type != 5:
                 colorObj.ConvertToRGB()
            return [colorObj.RGBRed,colorObj.RGBGreen,colorObj.RGBBlue]
        elif mode == 'CMYK':
            if colorObj.type != 2:
                 colorObj.ConvertToCMYK()
            return [colorObj.CMYKCyan,colorObj.CMYKMagenta,colorObj.CMYKYellow,colorObj.CMYKBlack]
        elif mode == 'CMY':
            if colorObj.type != 4:
                colorObj.ConvertToCMY()
            return [colorObj.CMYCyan,colorObj.CMYMagenta,colorObj.CMYYellow]
        elif mode == 'HSB':
            if colorObj.type != 6:
                 colorObj.ConvertToHSB()
            return [colorObj.HSBHue,colorObj.HSBBrightness,colorObj.HSBHue]
        elif mode == 'HLS':
            if colorObj.type != 7:
                 colorObj.ConvertToHLS()
            return [colorObj.HLSHue,colorObj.HLSLightness,colorObj.HLSSaturation]


    # 设置颜色值
    # colorObj 颜色对象
    # mode 返回的颜色模式
    # value 颜色值，数组格式
    def setColorValue(self,colorObj,value,mode="RGB"):
        if mode == 'RGB':
            if colorObj.type != 5:
                 colorObj.ConvertToRGB()
            return colorObj.RGBAssign(value[0],value[1],value[2])
        elif mode == 'CMYK':
            if colorObj.type != 2:
                 colorObj.ConvertToCMYK()
            return colorObj.CMYKAssign(value[0],value[1],value[2],value[3])
        elif mode == 'CMY':
            if colorObj.type != 4:
                colorObj.ConvertToCMY()
            return colorObj.CMYAssign(value[0],value[1],value[2])
        elif mode == 'HSB':
            if colorObj.type != 6:
                 colorObj.ConvertToHSB()
            return colorObj.HSBAssign(value[0],value[1],value[2])
        elif mode == 'HLS':
            if colorObj.type != 7:
                 colorObj.ConvertToHLS()
            return colorObj.HLSAssign(value[0],value[1],value[2])


    # 创建R颜色对象
    # createRGBColor([110,128,255],'t1')
    # name是作为搜索的一个key
    def createColorObj(self,value,name = '',mode="RGB"):
        colorObj = self.app.CreateColor()
        if name:
            colorObj.setname(name)
        self.setColorValue(colorObj,value,mode)
        return colorObj


    # 找到调色板对象
    def findPaletteObj(self,name):
        return self.app.PaletteManager.GetPalette(name)


    #转化对象
    def transformPaletteObjs(self, shapeObj):
        if self.getType(shapeObj) == 'str':
            return self.findPaletteObj(shapeObj)
        return shapeObj


    # 加载调色板
    def loadPalette(self,path):
        return self.app.Palettes.Open(path)


    # 创建调色板，如果没有就创建
    # nam 调色板名字
    # path 保存路径/默认文档路径
    # overwrite 是否覆盖，变成默认调色板，默认 不覆盖
    def accessPalette(self,name,path = '',overwrite = False):
        paletteObj = self.findPaletteObj(name)
        if paletteObj != None:
            return paletteObj
        # 默认保存文档路径
        if path == '':
            path = self.doc.filepath
        return self.app.Palettes.create(name,path,overwrite)


    # 删除调色板
    # nameObj 调色板名字/调色板对象
    def removePlette(self,nameObj):
        paletteObj = self.transformPaletteObjs(nameObj)
        if paletteObj == None:
            return
        self.setPletteDisabled(paletteObj)
        paletteObj.delete()


    # 设置调色板可用
    # nameObj 调色板名字/调色板对象
    def setPletteEnabled(self,nameObj):
        paletteObj = self.transformPaletteObjs(nameObj)
        paletteObj.Open()
        return paletteObj


    # 禁用调色板
    # nameObj 调色板名字/调色板对象
    def setPletteDisabled(self,nameObj):
        paletteObj = self.transformPaletteObjs(nameObj)
        paletteObj.Close()
        return paletteObj


    # 增加颜色到指定的调色板
    # nameObj 调色板名字/调色板对象
    # colorObj 颜色对象
    # index 增加指定的索引位置
    def addPletteColor(self,nameObj,colorObj,index=''):
        paletteObj = self.transformPaletteObjs(nameObj)
        #后追加
        if index == '':
            return paletteObj.addcolor(colorObj)
        else:
            if index >= paletteObj.ColorCount:
                index = paletteObj.ColorCount + 1
            # 指定插入的位置
            return paletteObj.InsertColor(index,colorObj)


    # 替换调色板颜色,通过索引
    # nameObj 调色板名字/调色板对象
    # colorObj 颜色对象
    # index 替换的索引
    def replacePletteColorByIndex(self,nameObj,colorObj,index):
        paletteObj = self.transformPaletteObjs(nameObj)
        self.removePletteColor(paletteObj,index)
        return self.addPletteColor(paletteObj,colorObj,index)


    # 替换调色板颜色,通过名字
    # nameObj 调色板名字/调色板对象
    # colorObj 颜色对象
    # name    颜色名字
    def replacePletteColorByName(self,nameObj,colorObj,name=''):
        paletteObj = self.transformPaletteObjs(nameObj)
        if name == '':
            name = colorObj.name
        colorIndex = paletteObj.findcolor(name)
        if colorIndex > 0:
            self.removePletteColor(paletteObj,colorIndex)
            return self.addPletteColor(paletteObj,colorObj,colorIndex)


    # 删除颜色
    # nameObj 调色板名字/调色板对象
    # index 需要删除的索引
    def removePletteColor(self,nameObj,index):
        paletteObj = self.transformPaletteObjs(nameObj)
        return paletteObj.RemoveColor(index)


    # 获取调色板指定颜色
    # nameObj 调色板名字/调色板对象
    # key     颜色对象关键字
    def getPaletteColor(self,nameObj,name=''):
        paletteObj = self.transformPaletteObjs(nameObj)
        if paletteObj == None:
            return
        colorIndex = paletteObj.findcolor(name)
        return paletteObj.Color(colorIndex)

