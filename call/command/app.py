import subprocess
import sys
import json

# obj = subprocess.Popen(cmdStr, stdin=subprocess.PIPE, stdout=subprocess.PIPE ,stderr=subprocess.PIPE)
# print(obj.stdin.write('ls\n'.encode('utf-8')))

python2json = {}
python2json["path"] = "C:\\Users\\Administrator\\Desktop\\cdr\\创意红色大气党建展板.cdr"
python2json["fontJson"] = "True"
json_str = json.dumps(python2json)

cmdStr = [
    "D:\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe", json_str]
child = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE,
                         stdin=subprocess.PIPE, stderr=subprocess.PIPE)
for line in child.stdout.readlines():
    output = line.decode('GBK')
    print(output)
