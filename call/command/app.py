import subprocess
import sys

# obj = subprocess.Popen(cmdStr, stdin=subprocess.PIPE, stdout=subprocess.PIPE ,stderr=subprocess.PIPE)
# print(obj.stdin.write('ls\n'.encode('utf-8')))

cmdStr = ["D:\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe",
          "C:\\Users\\Administrator\\Desktop\\cdr\\从严治党建文化墙.cdr","fontJson:true"]
child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('GBK')
    print(output)
