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
externalData = "{'address':'aaaaaaaaaaaa','url':''}"

cmdStr = ["D:\\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe",command,externalData]

child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('UTF-8')
    print(output)
