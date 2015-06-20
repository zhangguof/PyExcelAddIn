#-*- coding:utf-8 -*-

def test():
	os=__import__("os",globals(),fromlist=["*"])
	globals().update(os.__dict__)


test()
print getcwd()