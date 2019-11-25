from cdr import CDR

# 打开文档
def open():
    print(CDR('C:\\Users\\Administrator\\Desktop\\11.cdr'))


# 获页面内容
# 传递页面搜索
# 不传，默认获取所有页面数据
def getPageContent(pageIndex=""):
    print('返回', CDR().get(pageIndex))


def setPageContent(pageIndex="",path=""):
    data = {'logo': {'pageIndex': 1, 'value': 'C:\\Users\\Administrator\\Desktop\\QQ图片20191120100822.png'}, 'job': {'pageIndex': 1, 'value': '设计总监3'}, 'name': {'pageIndex': 1, 'value': 
'张天奕1'}, 'qrcode': {'pageIndex': 2, 'value': ''}, 'address': {'pageIndex': 2, 'value': '北京市朝阳区农展馆南路11114号\x0b瑞辰国际中心1807室1'}, 'mobile': {'pageIndex': 2, 'value': 
'888888881'}, 'phone': {'pageIndex': 2, 'value': ''}, 'url': {'pageIndex': 2, 'value': 'www.tianyishidai1.com'}, 'bjnews': {'pageIndex': 2, 'value': ''}, 'email': {'pageIndex': 2, 'value': '684755881@qq.com'}, 'qq': {'pageIndex': 2, 'value': ''}, 'company': {'overflow': True, 'pageIndex': 2, 'value': '北京天奕时代sdfsdf创意设计有限公12312312司1'}}
    CDR().set(data)

if __name__ == '__main__':
    # print( test.get("aaaa") ==None)
    # getPageContent()
    # open()
    setPageContent()