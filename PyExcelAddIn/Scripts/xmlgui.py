#-*- coding:utf-8 -*-

import clr,sys
#sys.path.append(r"    C:\Program Files (x86)\Microsoft Visual Studio 11.0\Visual Studio Tools for Office\PIA\Office14")

#office_dll = r"Office.dll"
#clr.AddReferenceToFile(office_dll)
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox

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

		def OnButtonClick(self, control):
			Id = control.Id
			MessageBox.Show("in click:"+str(control.Id))
			f = getattr(self.__class__,"on_button_%s"%Id,None)
			MessageBox.Show(str(f))

			if f:
				f(self)
			else:
				MessageBox.Show("Can't found Button(id:%s) click Event."%Id)

		def on_button_testbtn(self):
			MessageBox.Show("in test btn...")

		def on_button_test2(self):
			MessageBox.Show("in test 2")

		def Ribbon_Load(self,ribbon):
			self.ribbon = ribbon
	xml_rib = XmlRibbon()

	return xml_rib





