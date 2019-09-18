#coding=utf-8
#
try:
    print("test1")
    print("running")
    s = 1/0
except BaseException:
    print("baseException")
finally:
    print("finally")
