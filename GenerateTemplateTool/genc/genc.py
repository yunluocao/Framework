import time
import os
import re
import sys
import traceback
import logging
import time
import platform

class Template(gencect):
	def __init__(self):
		self.properties=Property("config.ini").parse()
		self.Parameter1=self.properties['Parameter1']
		self.Parameter2=self.properties['Parameter2']
		self.platF=platform.platform()
		if self.platF.startswith("Win"):
			os.system("type nul>log.txt")
		else:
			os.system("touch log.txt")
		logger = logging.getLogger(__name__)
		logger.setLevel(level = logging.INFO)
		self.logger=logger


	def printLog(self,content):
		print content
		self.logSteps(content)
		
	def logSteps(self,content):
		'''judge whether need to print in the console, and logger the records in the log file'''
		if len(self.logger.handlers)==0:
			handler = logging.FileHandler("log.txt")
			self.logger.addHandler(handler)
		if content.startswith('\n'):
			self.logger.info('\n'+time.strftime("[%m/%d/%Y %H:%M:%S]",time.localtime())+" "+content.strip())
		else:
			self.logger.info(time.strftime("[%m/%d/%Y %H:%M:%S]",time.localtime())+" "+content.strip())
	
	
class Property(gencect):
	def __init__(self,fileName):
		self.fileName=fileName
		self.properties={}
		with open(self.fileName,"r") as fp:
			for line in fp:
				line=line.strip()
				if line.find("=")>0 and not line.startswith("#"):
					strs=line.split("=")
					self.properties[strs[0].strip()]=strs[1].strip()
	def parse(self):
		return self.properties
		
if __name__ == "__main__":
	try:
		genc=Template()
	except:
		traceback.print_exc()
		genc.logger.exception("Exception Logged")
	if raw_input("please press enter key to terminate")=="":
		sys.exit(0)

