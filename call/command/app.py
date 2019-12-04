import subprocess
import sys
import json
import win32file
import win32api
import win32con
import win32com.client
from win32com.client import Dispatch, constants

# obj = subprocess.Popen(cmdStr, stdin=subprocess.PIPE, stdout=subprocess.PIPE ,stderr=subprocess.PIPE)
# print(obj.stdin.write('ls\n'.encode('utf-8')))


# 路径
# path = 'C:\\Users\\Administrator\\Desktop\\黑色高档服装名片设计.cdr'
path = 'C%3A%5CUsers%5CAdministrator%5CDocuments%5C%E7%A7%92%E7%A7%92%E5%AD%A6%5C08d865ff882eca28af9f01798f73cfe0%5C%E5%88%9B%E6%84%8F%E9%87%91%E8%89%B2%E7%A7%91%E6%8A%80%E5%95%86%E5%8A%A1%E5%90%8D%E7%89%87%5C%E5%88%9B%E6%84%8F%E9%87%91%E8%89%B2%E7%A7%91%E6%8A%80%E5%95%86%E5%8A%A1%E5%90%8D%E7%89%87.cdr'
# path = 'C:%5CUsers%5CAdministrator%5CDesktop%5C黑色高档服装名片设计.cdr'
# 配置
# open 打开文档 参数 (open,path)
# get:fontJson 获取字体json (get:fontJson,path)
# get:pageSize 获取页面尺寸 (get:pageSize,path)
# set:style 设置样式文件 (set:style,样式文件路径, 文档路径)

# 全部
# get:text 提取文本数据 (get:text,path)
# set:text 提取文本数据 (set:text,data,path)

# 指定页面
# get:text 获取页面尺寸 (get:text,page,path)
# set:text 提取文本数据 (set:text,data,page,path)

command = "set:font"

stylePath = "C:\\Users\\Administrator\\Desktop\\1618d6a4-e32c-11e9-b5e8-086266c80046.cdss"

# 外部数据
# externalData = "{'logo':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctt.jpg'}"

data = "{'job':{'pageIndex':'1','value':'设计总监'},'name':{'pageIndex':'1','value':'张天'},'address':{'pageIndex':'2','value':'北京市朝阳区农展馆南路13号\u000b瑞辰国际中心1807室'},'mobile':{'pageIndex':'2','value':'168-88888881'},'phone':{'pageIndex':'2','value':'0737-88888882111'},'url':{'pageIndex':'2','value':'www.tianyishidai.ccc'},'email':{'pageIndex':'2','value':'68475588@qq.com'}}"


data1 = "{'mobile':{'pageIndex':'1','value':'168-888888882'},'phone':{'pageIndex':'1','value':'123'},'url':{'pageIndex':'1','value':'68475588@qq.com\u000bwww.tianyishidai.com'},'bjnews':{'pageIndex':'1','value':''},'address':{'pageIndex':'1','value':'北京市朝阳区农展馆南路13号\u000b瑞辰国际中心1807'},'job':{'pageIndex':'1','value':'设计总监12'},'name':{'pageIndex':'1','value':'张天奕\u000b设计总监'},'company':{'pageIndex':'2','value':'北京天奕时代创意设计有限公司'},'qrcode':{'pageIndex':'2','value':''}}"

data2 = "{'mobile':{'pageIndex':'1','value':'168-11111'},'phone':{'pageIndex':'1','value':'2222'},'address':{'pageIndex':'1','value':'13北京市朝阳区农展馆\u000b南路13号瑞辰国际中心1807室'},'':{'pageIndex':'1','value':''},'email':{'pageIndex':'1','value':'68475588@qq.com\u000bwww.tianyishidai.com'},'qq':{'pageIndex':'1','value':'11111111111'},'job':{'pageIndex':'1','value':'设计总监'},'name':{'pageIndex':'1','value':'张天奕\u000b\u000b'},'company':{'pageIndex':'2','value':'北京天奕时代创意设计有限公司'},'logo':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctt.jpg'}}"

data3 = "{'logo':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctttt.jpg'},'qrcode':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctt.jpg'}}"

data4 = "{'logo':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctttt.jpg'}}"

data5 = "{'address':{'pageIndex':'2','value':'123测试4445'}}"

data6 = "{'email':{'pageIndex':'1','value':'68475588@qq.com'},'qq':{'pageIndex':'1','value':''},'address':{'pageIndex':'1','value':'北京市朝阳区农展馆南路13号\u000b瑞辰国际中心1807室'},'url':{'pageIndex':'1','value':'www.tianyishidai.com'},'bjnews':{'pageIndex':'1','value':''},'mobile':{'pageIndex':'1','value':'168-88888888'},'phone':{'pageIndex':'1','value':''},'job':{'pageIndex':'1','value':'职务/设计总监'},'name':{'pageIndex':'1','value':'张天奕'},'logo':{'pageIndex':'1','value':'C%3A%5CUsers%5CAdministrator%5CDesktop%5C111%5Chome.jpg'},'':{'pageIndex':'1','value':''},'qrcode':{'pageIndex':'2','value':'C%3A%5CUsers%5CAdministrator%5CDesktop%5C111%5Chome.jpg'}}"

data7 = "{'pageIndex':'1','value':'北京市朝阳区农展馆南路13号瑞辰国际中心1807室'},'mobile':{'pageIndex':'1','value':'168-888888888'},'phone':{'pageIndex':'1','value':''},'email':{'pageIndex':'1','value':'6847588@qq.com'},'qq':{'pageIndex':'1','value':''},'url':{'pageIndex':'1','value':'www.tianyishidai.com'},'bjnews':{'pageIndex':'1','value':''},'logo':{'pageIndex':'1','value':''},'job':{'pageIndex':'1','value':'设计总监'},'name':{'pageIndex':'1','value':'张天奕'},'companyname':{'pageIndex':'1','value':'Beijing\u200B\u200BTian\u200B\u200Byi\u200B\u200BTimes\u200B\u200BCreative\u200B\u200BDesign\u200B\u200BCo.,\u200B\u200BLtd.'},'company':{'pageIndex':'1','value':'北京天奕时代创意设计有限公司'},'qrcode':{'pageIndex':'2','value':''},'logo2':{'pageIndex':'2','value':''},'slogan':{'pageIndex':'2','value':'公司业务范围：\u000b平面设计/UI设计/3D效果图/新媒体运营'}}"

data8 = "{'mobile':{'pageIndex':'1','value':'16888888888'},'phone':{'pageIndex':'1','value':'11111111'},'address':{'pageIndex':'1','value':'北京市朝阳区农展馆南路13号瑞辰国际中心1807室'},'bjnews':{'pageIndex':'1','value':'168-88888888'},'url':{'pageIndex':'1','value':'11111111111'},'job':{'pageIndex':'1','value':'设计总监'},'name':{'pageIndex':'1','value':'张天奕1'}}"

data9 = "{'qrcode':{'pageIndex':'1','value':'C%3A%5CUsers%5CAdministrator%5CDesktop%5C111%5Chome.jpg'},'address':{'pageIndex':'1','value':'北京市朝阳区农展馆南路13号\r瑞辰国际中心1807室'},'mobile':{'pageIndex':'1','value':'168-88888888'},'phone':{'pageIndex':'1','value':'1'},'email':{'pageIndex':'1','value':'6847588@qq.com'},'qq':{'pageIndex':'1','value':'11111111'},'bjnews':{'pageIndex':'1','value':'16888888888'},'url':{'pageIndex':'1','value':'11111'},'job':{'pageIndex':'1','value':'设计总监'},'name':{'pageIndex':'1','value':'张天奕'},'logo':{'pageIndex':'2','value':'C%3A%5CUsers%5CAdministrator%5CDesktop%5C111%5C%EF%BC%91.jpg'},'companyname':{'pageIndex':'2','value':'www.tianyishidai.com'}}"

data10 = "{'logo':{'pageIndex':'1','value':'C%3A%5CUsers%5CAdministrator%5CDesktop%5C111%5C1.png'}}"

cmdStr = [
    "D:\\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe", 'set:text', data10]

child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('UTF-8')

    print(output)
