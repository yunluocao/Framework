import os 
import re
import collections
def ContentCompare(List1,List2):
	List1=sorted(List1)
	List2=sorted(List2)
	ListUnFound=[]
	if List1==List2:
		return ListUnFound
	for item1 in List1:
		Flag=False
		if item1 in List2:
			Flag=True
		else:
			for item2 in List2:
				if re.search(r"/$",item1) and not re.search(r"/$",item2):
					if item1 in item2:
						Flag=True
						break
				elif re.search(r"/$",item2) and not re.search(r"/$",item1):
					if item2 in item1:
						Flag=True
						break
		if not Flag:
			ListUnFound.append(item1)
	return ListUnFound



ExCurrentPath=os.path.join(os.getcwd(),"Expected")
AcCurrentPath=os.path.join(os.getcwd(),"Actual")
FileResult=os.path.join(os.getcwd(),"Result")

os.chdir(ExCurrentPath)
ExpectedFileList=os.listdir(os.getcwd())
os.chdir(AcCurrentPath)
ActualFileList=os.listdir(os.getcwd())
ActualDict=collections.OrderedDict()
ExpectDict=collections.OrderedDict()
for acFile in ActualFileList:
	if re.search("from",acFile):
		ActualDict[acFile]=re.findall(r'from (.*).log',acFile)[0]
	elif re.search("to",acFile):
		ActualDict[acFile]=re.findall(r'to (.*).log',acFile)[0]
for exFile in ExpectedFileList:
	ExpectDict[re.findall(r'(.*)-.*.txt',exFile)[0]]=exFile
	
# print ActualDict
# print ExpectDict
CISPath=r"C:\\clover\\host\\cis6.1\\integrator\\"

ActualAddFileList=[]
ActualLackFileList=[]
for acFile in ActualFileList:
	fileExpected =os.path.join(ExCurrentPath,ExpectDict[ActualDict[acFile]])
	fileActual=os.path.join(AcCurrentPath,acFile)
	ListExpected=[]
	ListActual=[]
	with open(fileActual,"r") as fpac:
		for i in fpac:
			if re.match(CISPath,i):
				ListActual.append(re.sub(CISPath,"",i.strip()).replace("\\","/"))
	
	with open(fileExpected,"r") as fpex:
		for i in fpex:
			ListExpected.append(i.strip())
	ActualLackFileList.append("Expected File %s couldn't find in the actual file %s:" %(ExpectDict[ActualDict[acFile]],acFile))
	ActualLackFileList.append(ContentCompare(ListExpected,ListActual))
	ActualAddFileList.append("Actual File %s couldn't find in the expected file %s:" %(acFile,ExpectDict[ActualDict[acFile]]))
	ActualAddFileList.append(ContentCompare(ListActual,ListExpected))
	
	
	
with open(FileResult,"w") as fp:
	for i in ActualLackFileList:
		fp.write(str(i)+'\n\n')
	fp.write('\n')
	for i in ActualAddFileList:
		fp.write(str(i)+'\n\n')
