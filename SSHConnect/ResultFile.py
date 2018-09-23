#encoding=utf-8

import re
import sys
import os
def findDateStatus(DateVal):
	RootWarTestingPath=r"C:\cloverleaf\csc6.3\RootWarTesting"
	fp=open(os.path.join(RootWarTestingPath,"Result.txt"),"r")
	content=fp.readlines()
	fp.close()
	result=re.findall(r"(.*%s.*) log (.*)\n"%DateVal,"".join(content))
	status=result[0][1]
	print "The status is %s\n"%status
	filePath=os.path.join(RootWarTestingPath,"log","log%s.txt"%DateVal)
	with open(filePath,"r") as fp:
		for line in fp:
			print line,'\n'
	
	
	
	

if __name__=="__main__":
	findDateStatus(sys.argv[1])