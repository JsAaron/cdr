import utils
import result
import globalData

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

    def __setAllField(self, shape):
        key = utils.getKeyEnglish(shape.Name)
        if key:
            result.saveInputFiled(key)

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
        if self.field_4 == True:
            if globalData.hasValue("bjnews") or globalData.hasValue("url"):
                self.visibleField = "4字段"

        # 如果有3字段
        if self.field_3 == True:
            if globalData.hasValue("email") or globalData.hasValue("qq"):
                self.visibleField = "3字段"

    # 初始化

    def initField(self, key, shapes, pageIndex):
        self.__setField(key)
        self.__setVisibleField()
        self.__proccessText(shapes, pageIndex)

    # '判断是否需要合并的数据

    def getRangeScope(self,key):
        if key == "url" or key == "bjnews" or key == "mobile" or key == "phone" or key == "email" or key == "qq":
            return True
        else:
            return False
