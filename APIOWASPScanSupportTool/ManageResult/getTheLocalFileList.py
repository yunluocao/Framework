#encoding=utf-8
import os
import json

List=os.listdir(os.getcwd())
NewList=filter(lambda x:"." not in x,List)

jsObj=json.dumps(NewList)
with open("FileList.txt","w") as fp:
	fp.write(jsObj)

