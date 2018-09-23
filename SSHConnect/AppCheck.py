#encoding=utf-8
import os
os.system("rm -rf taskList.txt")
os.system("tasklist>taskList.txt")
with open ("taskList.txt","r") as fp:
	for line in fp:
		if "RootWarTesting.exe" in line:
			break
	else:
		print "The RootWarTesting.exe isn't executed in the remote"