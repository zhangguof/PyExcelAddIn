#-*- coding:utf-8 -*-

import clr,sys
clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import *

ExcelAddIn = None

def init(namespace,ScriptPath,Libs):
	global ExcelAddIn
	clr.AddReference(namespace)
	ExcelAddIn = __import__(namespace,globals(),fromlist=["*"])
	sys.path.append(ScriptPath)
	sys.path.append(Libs)

	import conf

	conf.ScriptPath = ScriptPath
	conf.ExcelAddIn = ExcelAddIn




def test(name):
	return "You name is :%s"%name

def getRibbon():
	import xmlgui
	return xmlgui.get_xml_ribbon(ExcelAddIn.IPyRibbon)
