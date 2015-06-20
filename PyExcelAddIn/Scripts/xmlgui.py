#-*- coding:utf-8 -*-

import clr,sys
#sys.path.append(r"    C:\Program Files (x86)\Microsoft Visual Studio 11.0\Visual Studio Tools for Office\PIA\Office14")

#office_dll = r"Office.dll"
#clr.AddReferenceToFile(office_dll)
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import *

#from Microsoft.Office.Core import IRibbonExtensibility

import conf
import os


def get_xml_ribbon(IpyRibClass):
	class XmlRibbon(IpyRibClass):
		# def __init__(self):
		# 	self.ribbon = ribbon
		def GetCustomUI(self, ribbonId):
			with open(os.path.join(conf.ScriptPath,conf.gui_xml_file),"rb") as f:
				xml_str = f.read() 
			#MessageBox.Show(ribbonId)
			return xml_str

		def OnTest(self,*args):
			MessageBox.Show(str(args))

		def Ribbon_Load(self,ribbon):
			#self.ribbon = ribbon
			pass
	xml_rib = XmlRibbon()
	MessageBox.Show(str(IpyRibClass.__dict__))

	return xml_rib





