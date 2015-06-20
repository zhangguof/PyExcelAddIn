#-*- coding:utf-8 -*-

import clr,sys
clr.AddReference("System.Windows.Forms")
#clr.AddReference("Microsoft.Office.Tools.Ribbon")


from System.Windows.Forms import *
#from Microsoft.Office.Tools.Ribbon import RibbonBase


ExcelAddIn = None



def init(namespace,ScriptPath,Libs):
	global ExcelAddIn
	clr.AddReference(namespace)
	ExcelAddIn = __import__(namespace,globals(),fromlist=["*"])
	sys.path.append(ScriptPath)
	sys.path.append(Libs)
	

	class Ribbon1(ExcelAddIn.Ribbon):
		def __init__(self):
			self.Name = "TestRibbon"
			self.RibbonType = "Microsoft.Excel.Workbook"
	rib = Ribbon1()
	MessageBox.Show(str(rib.__dict__))

	



def test(name):
	return "You name is :%s"%name




