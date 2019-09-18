#coding=utf-8
try:
    print("test1")
    print("running 异常")
    s = 1/0
except BaseException:
    print("baseException 异常")
finally:
    print("finally")
