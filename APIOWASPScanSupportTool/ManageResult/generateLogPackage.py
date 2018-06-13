#encoding=utf-8
import os
import shutil
import zipfile
import re
import sys
import json

class moveFiles(object):
	'''to move the html file and package the active file into the related folder, it also could check the empty folder'''
	"to avoid error: non-default argument follows default argument, please put the default argument in the end"
	def __init__(self,path,MoveFile=True):
		self.MoveFile=MoveFile
		self.path=path
		self.main()
		
		
	def main(self):
		if os.path.exists("report"):
			shutil.rmtree("report")
			shutil.rmtree("session")
		self.BaseLocation=self.path
		List=self.findEmptyFolder()
		if List!=[]:
			print "there is empty folder, please refer the list %s" %List
			sys.exit(1)
		os.chdir(self.BaseLocation)
		os.mkdir("report")
		os.mkdir("session")
		self.generateThePassedList()
		if self.MoveFile:
			self.moveFilesMethod()
			self.packageActive()
			print "*"*30
			print "you could go to the report folder find the report, and go to the session folder find the active.zip package"
		
		
	def findEmptyFolder(self):
		List=[]
		for root, dirs, files in os.walk(self.BaseLocation):
			if dirs==[]:
				os.chdir(root)
				os.chdir("..")
				for i in os.listdir(os.getcwd()):
					if '.' not in i:
						if os.listdir(os.path.join(os.getcwd(),i))==[] and os.getcwd()+'\\'+i not in List:
							List.append(os.getcwd()+'\\'+i)
		return List
	
	def moveFilesMethod(self):
		for root, dirs, files in os.walk(self.BaseLocation):
			for file in files:
				bare_name = os.path.splitext(file)[0]
				extension_name = os.path.splitext(file)[1]
				if extension_name == '.html' and "report" not in root:
					print "start to move the the suite %s files" %root
					os.chdir(root)
					ActiveName=root.split("\\")[-1]
					print ActiveName,self.BaseLocation
					FolderName=re.match(r"%s\\(.*)\\%s" %(self.BaseLocation.replace("\\","\\\\"),ActiveName.replace("\\","\\\\")),root).group(1)
					ReportDestination=os.path.join(self.BaseLocation,"report",FolderName)
					try:
						os.makedirs(ReportDestination)
					except Exception:
						pass
					SessionDestination=os.path.join(self.BaseLocation,"session",FolderName,root.split("\\")[-1])
					try:
						os.makedirs(SessionDestination)
					except Exception:
						pass
					shutil.copy(file,ReportDestination)
					for root, dirs, files in os.walk(os.getcwd()):
						for file in files:
							if "html" not in file:
								shutil.copy(file,SessionDestination)
	
	def packageActive(self):
		os.chdir(os.path.join(self.BaseLocation,"session"))
		for root, dirs, files in os.walk(os.getcwd()):
			if 'active' in root:
				ActiveName=root.split("\\")[-1]
				os.chdir(root)
				os.chdir("..")
				#zip定义的地方就是生成zip的地方
				z=zipfile.ZipFile(ActiveName+".zip","w",zipfile.ZIP_DEFLATED,allowZip64=True)
				os.chdir(root)
				print "start to package the %s" %root
				# shutil.make_archive(ActiveName+"3", 'zip',os.getcwd())
				for file in os.listdir(os.getcwd()):
					z.write(file)
				z.close()
				os.chdir("..")
				shutil.rmtree(ActiveName)
				
	def generateThePassedList(self):
		List=os.listdir(self.BaseLocation)
		NewList=filter(lambda x:"." not in x,List)
		jsObj=json.dumps(NewList)
		with open(os.path.join(self.BaseLocation,"FileList.txt"),"w") as fp:
			fp.write(jsObj)
		print "the suite filelist is generated, please go to the below path %s to find" %(os.path.join(self.BaseLocation,"FileList.txt"))


		
		
		
if __name__=="__main__":
	moveObj=moveFiles(sys.argv[1],True)
	
