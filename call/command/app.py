import subprocess
import sys
import json

# obj = subprocess.Popen(cmdStr, stdin=subprocess.PIPE, stdout=subprocess.PIPE ,stderr=subprocess.PIPE)
# print(obj.stdin.write('ls\n'.encode('utf-8')))

# 路径
path = 'C:\\Users\\Administrator\\Desktop\\黄蓝-黑.cdr'
# 配置
# open 打开文档 参数 (open,path)
# get:pageSize 获取页面尺寸 (get:pageSize,path)
# get:fontJson 获取字体json (get:fontJson,path)
# get:text 提取文本数据（名片）(get:text,path)
# set:text 提取文本数据（名片）(set:text,externalData,path)
# set:style 设置样式文件 (set:style,样式文件路径, 文档路径)
command = "set:text"

stylePath = "C:\\Users\\Administrator\\Desktop\\1618d6a4-e32c-11e9-b5e8-086266c80046.cdss"

# 外部数据
# externalData = "{'logo':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctt.jpg'}"

data = "{'job':{'pageIndex':'1','value':'设计总监1'},'name':{'pageIndex':'1','value':'张天'},'address':{'pageIndex':'2','value':'北京市朝阳区农展馆南路13号\u000b瑞辰国际中心1807室'},'mobile':{'pageIndex':'2','value':'168-88888881'},'phone':{'pageIndex':'2','value':'0737-88888882'},'url':{'pageIndex':'2','value':'www.tianyishidai.com'},'email':{'pageIndex':'2','value':'68475588@qq.com'}}"

data1 = "{'mobile':{'pageIndex':'1','value':'168-888888882'},'phone':{'pageIndex':'1','value':'123'},'url':{'pageIndex':'1','value':'68475588@qq.com\u000bwww.tianyishidai.com'},'bjnews':{'pageIndex':'1','value':''},'address':{'pageIndex':'1','value':'北京市朝阳区农展馆南路13号\u000b瑞辰国际中心1807'},'job':{'pageIndex':'1','value':'设计总监12'},'name':{'pageIndex':'1','value':'张天奕\u000b设计总监'},'company':{'pageIndex':'2','value':'北京天奕时代创意设计有限公司'},'qrcode':{'pageIndex':'2','value':''}}"

data2 = "{'mobile':{'pageIndex':'1','value':'168-11111'},'phone':{'pageIndex':'1','value':'2222'},'address':{'pageIndex':'1','value':'13北京市朝阳区农展馆\u000b南路13号瑞辰国际中心1807室'},'':{'pageIndex':'1','value':''},'email':{'pageIndex':'1','value':'68475588@qq.com\u000bwww.tianyishidai.com'},'qq':{'pageIndex':'1','value':'11111111111'},'job':{'pageIndex':'1','value':'设计总监'},'name':{'pageIndex':'1','value':'张天奕\u000b\u000b'},'company':{'pageIndex':'2','value':'北京天奕时代创意设计有限公司'},'logo':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctt.jpg'}}"

data3 = "{'logo':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctttt.jpg'},'qrcode':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctt.jpg'}}"

data4 = "{'logo':{'pageIndex':'2','value':'C:%5CUsers%5CAdministrator%5CDesktop%5Ctttt.jpg'}}"

data5 = "{'address':{'pageIndex':'2','value':'123测试4445'}}"



cmdStr = [
    "D:\\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe", command,data5]

child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('UTF-8')

    print(output)
