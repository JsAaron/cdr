import subprocess
import sys


cmdStr = ["D:\\github\\cdr\\ConsoleApp\\ConsoleApp\\bin\\Debug\\ConsoleApp.exe",
          "C:\\Users\\Administrator\\Desktop\\cdr\\19规范版横版.cdr", "true"]
pi = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE)


lines = pi.stdout.readlines()
print(lines)
