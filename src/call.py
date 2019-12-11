from cdr import CDR

# 打开文档
def open():
    print(CDR('C:\\Users\\Administrator\\Desktop\\11.cdr'))

# 获页面内容
# 传递页面搜索
# 不传，默认获取所有页面数据
def getContent(pageIndex=""):
    print('返回', CDR().get(pageIndex))


def setContent(pageIndex="",path=""):
    data = {'logo': {'pageIndex': 1, 'value': 'C:\\Users\\Administrator\\Desktop\\111\\1.png'}}
    CDR().set(data)


def togglePage():
    print(CDR().togglePage(2))
    


if __name__ == '__main__':
    # togglePage()
    # print( test.get("aaaa") ==None)
    # getContent()
    # open()
    # setContent()
    CDR().drawDecorationTriangle("test","255, 0, 0",{"top":100,"left":200},'lefttop')

