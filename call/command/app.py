import subprocess
import sys
import json

# obj = subprocess.Popen(cmdStr, stdin=subprocess.PIPE, stdout=subprocess.PIPE ,stderr=subprocess.PIPE)
# print(obj.stdin.write('ls\n'.encode('utf-8')))

# 路径
path ='C:\\Users\\Administrator\\Desktop\\黄蓝-黑.cdr'
# 配置 
# open 打开文档 参数 (open,path)
# get:pageSize 获取页面尺寸 (get:pageSize,path)
# get:fontJson 获取字体json (get:fontJson,path)
# get:text 提取文本数据（名片）(get:text,path)  
# set:text 提取文本数据（名片）(set:text,externalData,path)  
command = "set:extract"
# 外部数据
externalData = "{'job':'老大','name':'沉稳'}"

cmdStr = [
    "D:\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe", "get:text"]

child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('UTF-8')
    print(output)

