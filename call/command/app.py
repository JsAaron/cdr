import subprocess
import sys
import json

# obj = subprocess.Popen(cmdStr, stdin=subprocess.PIPE, stdout=subprocess.PIPE ,stderr=subprocess.PIPE)
# print(obj.stdin.write('ls\n'.encode('utf-8')))

# 路径
path ='C:\\Users\\Administrator\\Desktop\\黄蓝-黑.cdr'

# 配置 
# pagesize 获取页面尺寸
# fontJson 启动字体json
# extract 提取文本数据（名片） 
config = "{'pagesize':'True','fontjson':'False','extract':'True'}"

# 外部数据
externalData = "{'test':{'aa':'11111111','bb':'11111111','cc':'11111111'}}"


cmdStr = [
    "D:\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe", path,config,externalData]

child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('GBK')
    print(output)

