from cdr import CDR

#打开文档
def open():
    print(CDR('C:\\Users\\Administrator\\Desktop\\11.cdr'))


#获页面内容
def getPageContent():
    print(CDR().getPageContent(1))


if __name__ == '__main__':
    getPageContent()
    # open()
