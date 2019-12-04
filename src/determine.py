import utils
import result
import prarm

# 数据判断类，是否分行


class Determine():

    def __init__(self):
        # 模板存在的字段
        self.field_2 = False
        self.field_3 = False
        self.field_4 = False
        self.visibleField = "2字段"

        # 组合字段状态定义
        self.cdr_url = False
        self.cdr_bjnews = False
        self.cdr_mobile = False
        self.cdr_phone = False
        self.cdr_email = False
        self.cdr_qq = False

    def __setField(self, key):
        if key == "2字段":
            self.field_2 = True
        elif key == "3字段":
            self.field_3 = True
        elif key == "4字段":
            self.field_4 = True

    # 初始化字段的状态
    # 涉及到状态合并的问题处理

    def __setState(self, key):
        if key == "url":
            self.cdr_url = True
        elif key == "bjnews":
            self.cdr_bjnews = True
        elif key == "mobile":
            self.cdr_mobile = True
        elif key == "phone":
            self.cdr_phone = True
        elif key == "email":
            self.cdr_email = True
        elif key == "qq":
            self.cdr_qq = True

    # 默认保存所有字段
    def __setAllField(self, shape):
        key = utils.getKeyEnglish(shape.Name)
        if key:
            result.saveInputFiled(key)

    # 文本预处理
    def __proccessText(self, shapes, pageIndex):
        cdrTextShape = 6
        cdrGroupShape = 7
        for tempShape in shapes:

            # 保存所有字段
            self.__setAllField(tempShape)

            # 组
            if tempShape.Type == cdrGroupShape:
                self.__proccessText(tempShape.Shapes, pageIndex)

            # 文本
            if tempShape.Type == cdrTextShape:
                self.__setState(utils.getKeyEnglish(tempShape.Name))

    # 设置使用层级模板
    def __setVisibleField(self):
        # 如果有4字段 显示层级4
        if self.field_4:
            if prarm.hasValue("bjnews") or prarm.hasValue("url"):
                self.visibleField = "4字段"

        # 如果有3字段
        if self.field_3:
            if prarm.hasValue("email") or prarm.hasValue("qq"):
                self.visibleField = "3字段"

    # =======================对外接口=======================

    # 初始化
    def initField(self, key, shapes, pageIndex):
        self.__setField(key)
        self.__setVisibleField()
        self.__proccessText(shapes, pageIndex)

    # 判断是否需要合并的数据
    def getRangeScope(self, key):
        if key == "url" or key == "bjnews" or key == "mobile" or key == "phone" or key == "email" or key == "qq":
            return True
        else:
            return False

    def getVisibleField(self):
        return self.visibleField

    def getMergeValue(self, key):
        newValue = prarm.getExternalValue(key)

        if key == "url":
            # url + bjnews
            if prarm.hasValue("bjnews") and self.cdr_bjnews == False:
                user_bjnews = prarm.getExternalValue("bjnews")
                newValue = newValue + '\n' + user_bjnews
        elif key == "bjnews":
             # url + bjnews
            if prarm.hasValue("url") and self.cdr_url == False:
                user_url = prarm.getExternalValue("url")
                newValue = user_url + '\n' + newValue
        elif key == "mobile":
             # 没有电话字段，但是用户设置了手机
            if prarm.hasValue("phone") and self.cdr_phone == False:
                user_phone = prarm.getExternalValue("phone")
                newValue = newValue + '\n' + user_phone
        elif key == "phone":
             # 没有手机字段，但是用户设置了电话
            if prarm.hasValue("mobile") and self.cdr_mobile == False:
                user_mobile = prarm.getExternalValue("mobile")
                newValue = user_mobile + '\n' + newValue
        elif key == "email":
             # 邮箱/QQ
            if prarm.hasValue("qq") and self.cdr_qq == False:
                user_qq = prarm.getExternalValue("qq")
                newValue = newValue + '\n' + user_qq
        elif key == "qq":
             # email + qq
            if prarm.hasValue("email") and self.cdr_email == False:
                user_email = prarm.getExternalValue("email")
                newValue = user_email + '\n' + newValue

        return newValue

    def __setVisible(self,activeLayer, name, visibleLayerName):
        if name == visibleLayerName:
            activeLayer.Visible = True
        else:
            activeLayer.Visible = False

    # 设置层级的可见性
    # 如果网址/公众号，都没有，那么要隐藏“4 字段”图层，显示“3 字段”图层。如果邮箱/QQ 号，也没有，那么就显示“2 字段图层
    def setLayerVisible(self,activeLayer, visibleLayerName):
        name = activeLayer.Name
        if name == "2字段":
            self.__setVisible(activeLayer, name, visibleLayerName)
        elif name == "3字段":
            self.__setVisible(activeLayer, name, visibleLayerName)
        elif name == "4字段":
            self.__setVisible(activeLayer, name, visibleLayerName)
