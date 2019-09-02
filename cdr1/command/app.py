import subprocess
import sys

cmdStr = ["D:\\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe","C:\\Users\\Administrator\\Desktop\\test.cdr"]
pi = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE)


lines = pi.stdout.readlines()
print(lines)
