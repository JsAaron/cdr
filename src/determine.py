import utils

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
            self.field_ = True

    def __setAllField(self,shape):
        key = utils.getKeyEnglish(shape.Name)
        print(key)

    # 文本预处理
    def __proccessText(self, shapes, pageIndex):
        cdrShapeType = 6
        cdrShapeType = 7
        for shape in shapes:
            self.__setAllField(shape)

    # 初始化
    def initField(self, key, shapes, pageIndex):
        self.__setField(key)
        # self.__setVisibleField()
        self.__proccessText(shapes, pageIndex)
