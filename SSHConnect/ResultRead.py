#encoding=utf-8

import re
import sys
import os

class fileRead(object):
	def __init__(self,DateVal,readFlag):
		# self.RootWarTestingPath=r"C:\cloverleaf\csc6.3\RootWarTesting"
		self.RootWarTestingPath=r"C:\APythonTest\AFramework\SSHConnect"
		self.resultFilePath=os.path.join(self.RootWarTestingPath,"Result.txt")
		self.DateVal=DateVal
		self.readFlag=readFlag
		
	def readLogFile(self,result):
		if self.readFlag.upper()=="Y":
			for i in result:
				filePath=os.path.join(self.RootWarTestingPath,"log","log%s.txt"%i[0])
				print "start to read the file %s"%filePath
				with open(filePath,"r") as fp:
					for line in fp:
						print line,'\n'
		
		
		
	def findDateStatus(self):
		fp=open(self.resultFilePath,"r")
		content=fp.readlines()
		print "The result is\n"
		if self.DateVal.upper()!="ALL":
			result=re.findall(r"(.*%s.*) log (.*)\n?"%self.DateVal,"".join(content))
		else:
			result=re.findall(r"(.*) log (.*)\n?","".join(content))
		print result,'\n'
		fp.close()
		return result

if __name__=="__main__":
	fR=fileRead(sys.argv[1],sys.argv[2])
	result=fR.findDateStatus()
	fR.readLogFile(result)

