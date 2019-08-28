import os
import sys

main = "C:\\Users\\Administrator\\source\\repos\\app\\app\\bin\\Debug\\app.exe"
r_v = os.system(main)

if sys.argv:
  print(sys.argv[1])


