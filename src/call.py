from cdr import CDR
test = {
    "aaa":1
}



#打开文档
def open():
    print(CDR('C:\\Users\\Administrator\\Desktop\\11.cdr'))


#获页面内容
def getPageContent():
    print('返回',CDR().getPageContent(2))


if __name__ == '__main__':
    # print( test.get("aaaa") ==None)

    getPageContent()
    # open()
