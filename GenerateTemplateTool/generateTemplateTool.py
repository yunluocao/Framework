#encoding=utf-8
import os
class generateTemTool(object):
	def __init__(self):
		pass
		
	def gen(self):
		fp=open("Template.ini","r")
		confContent=fp.readlines()
		fp.close()
		content=""
		for line in confContent:
			if not line.startswith("ToolName"):
				propertyName=line.split("=")[0]
				if "||False" not in line:
					content+="\t\tself.%s=eval(self.properties['%s'])\n"%(propertyName,propertyName)
				else:
					content+="\t\tself.%s=self.properties['%s']\n"%(propertyName,propertyName)
			else:
				Name=line.split("=")[1].strip()
		if not os.path.exists(Name):
			os.system("mkdir %s"%Name)
		
		with open("%s\config.ini"%Name,"w") as fp:
			fp.write("".join(confContent).replace("ToolName=%s\n"%Name,"").replace("||False",""))
		fp=open("Template.py","r")
		scriptContent=fp.readlines()
		fp.close()
		
		with open("%s\%s.py"%(Name,Name),"w") as fp:
			fp.write("".join(scriptContent).replace('self.properties=Property("config.ini").parse()\n','self.properties=Property("config.ini").parse()\n'+content).replace("obj",Name))
		

if __name__ == "__main__":
	obj=generateTemTool()
	obj.gen()