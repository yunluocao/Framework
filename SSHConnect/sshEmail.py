#encoding=utf-8
import win32com.client as win32
import warnings
import sys
import base64
reload(sys)
sys.setdefaultencoding('utf8')
warnings.filterwarnings('ignore')
import os
import ssh
import schedule
import datetime
import re
import schedule
import time
import traceback

class sshEmail(object):
	
	def __init__(self):
		self.properties=Property("sshEmailCon.ini").parse()
		self.remoteMachine=self.properties['remoteMachine']
		self.remoteUser=self.properties['remoteUser']
		self.remotePwd=self.properties['remotePwd']
		self.outLookPath=self.properties['outLookPath']
		self.attachFile=self.properties['attachFile']
		self.receivers=eval(self.properties['receivers'])
		self.receivers_debug=eval(self.properties['receivers_debug'])
		self.inboxFolderIndex=6
		self.outlook=self.outLookInit()
		self.submitTime=self.dateGet()
		self.status,self.content,self.errorContent=self.executionStatus()
		self.sub=self.submitTime+" testing result is %s"%self.status
		self.body=self.bodyValue()
		self.outLookCheckNum=1
	def dateGet(self):
		now_time = datetime.datetime.now()
		yes_time = now_time + datetime.timedelta(days=-1)
		yes_time_nyr = yes_time.strftime('%Y%m%d')
		return yes_time_nyr
		
	def bodyValue(self):
		if self.status=="passed":
			body="Hi All:\n\n" \
		 + "The date:%s root war basic testing is pass\n"%self.submitTime \
		 + "Please refer the build package(ROOT.war) is available at ftp://usspvlcentos64/Platform_Releases/CSC/Web_GUI_Daily/%s/"%self.submitTime +"\n" \
		 + "\n\nThanks\nBest Regards\n" \
		 + "Jonson"
		else:
			body="Hi All:\n\n" \
		 + "The date:%s root war basic testing is failed\n"%self.submitTime \
		 + "It failed by the below issue:\n" \
		 + '\n'+self.errorContent+'\n' \
		 + "\n\nThanks\nBest Regards\n" \
		 + "Jonson"
		return body
	
	def sshConnect(self,app):
		# New SSHClient
		client = ssh.SSHClient()
		# Default accept unknown keys
		client.set_missing_host_key_policy(ssh.AutoAddPolicy())
		# Connect
		client.connect(self.remoteMachine, port=22, username=self.remoteUser, password=self.remotePwd)
		# Execute shell remotely
		stdin, stdout, stderr = client.exec_command("%s %s"%(app,self.submitTime))
		return stdout.read().strip()
		
		
	def executionStatus(self):
		stdContent=self.sshConnect("ResultFile.py")
		status=re.search(r"The status is (.*)\n",stdContent).group(1).strip()
		content=stdContent.strip(r"The status is %s\n"%status).strip()
		errorContent=re.search(r"Exception: (.*)\n?",content).group(1).strip()
		return status,content,errorContent
		
	def outLookInit(self):
		try:
			outlook = win32.Dispatch('outlook.application')
			#actually the application is not opened by the script,it would have the below exception:
			#pywintypes.com_error: (-2146959355, 'Server execution failed', None, None)
			#it's need to closed,then open it could use it.
		except Exception,e:
			outlook=self.restartOutlook()
		return outlook
		
	def restartOutlook(self):
		os.system("TASKKILL /F /IM outlook.exe /T")
		outlook = win32.Dispatch('outlook.application')
		currentPath=os.getcwd()
		os.chdir(self.outLookPath)
		os.system("OUTLOOK.EXE")
		os.chdir(currentPath)
		return outlook
		
	def sendemail(self,receivers,sub,body,attachFile=None):
		# os.system("TASKKILL /F /IM outlook.exe /T")
		mail = self.outlook.CreateItem(0)
		mail.To = receivers[0]
		mail.Subject = sub.decode('utf-8')
		mail.Body = body.decode('utf-8')
		# 添加附件
		if attachFile:
			mail.Attachments.Add(attachFile)
		mail.Send()
		
	
	def mailLoopContentRead(self,Flag,count,messages,rangeValue,sub):
		time.sleep(60)
		for i in rangeValue:
			subject = messages[i].Subject
			if sub in subject:
				Flag=True
				print 'the mail subject \"%s\" is delivered successfully'%sub
				break
		else:
			if count<self.outLookCheckNum:
				self.outlook=self.restartOutlook()
				count+=1
			elif count==self.outLookCheckNum:
				raise Exception("the mail content %s couldn't be found"%sub)
		return Flag,count
	
		
		
		
	def mail_read(self,folderIndex,sub):
		try:
			count=0
			Flag=False
			while count<=self.outLookCheckNum and not Flag:
				mail = self.outlook.GetNamespace("MAPI").GetDefaultFolder(folderIndex)
				messages = mail.Items
				if len(messages)>=10:
					Flag,count=self.mailLoopContentRead(Flag,count,messages,range(len(messages)-1,len(messages)-20,-1),sub)
					
				else:
					Flag,count=self.mailLoopContentRead(Flag,count,messages,range(len(messages)),sub)
		except Exception,e:
			traceback.print_exc()
			

class Property(object):
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

def statusEmail():
	sE=sshEmail()
	sE.sendemail(sE.receivers,sE.sub,sE.body)
	sE.mail_read(sE.inboxFolderIndex,sE.sub)
	if sE.status=="failed":
		sE.sendemail(sE.receivers_debug,sE.sub+" details",sE.content)
		sE.mail_read(sE.inboxFolderIndex,sE.sub+" details")
	
	
def statusEmail_debug():
	sE=sshEmail()
	# sE.sendemail(sE.receivers_debug,sE.sub,sE.body)
	# sE.mail_read(sE.inboxFolderIndex,sE.sub)
	if sE.status=="failed":
		attachFile=screenShotEmail(sE)
		print attachFile
		sE.sendemail(sE.receivers_debug,sE.sub+" details",sE.content,attachFile)
		sE.mail_read(sE.inboxFolderIndex,sE.sub+" details")
		
		
def appExistEmail():
	sE=sshEmail()
	result=sE.sshConnect("AppCheck.py")
	if result!="":
		sE.sendemail(sE.receivers_debug,result,"")
		sE.mail_read(sE.inboxFolderIndex,result)
		
	
def screenShotEmail(sEobject):
	result=sEobject.sshConnect("ScreenShotFile.py")
	if result!="False":
		str=result
		file_str=open('issue.jpg','wb')
		file_str.write(base64.b64decode(str))
		file_str.close()
		attachFile=os.path.join(os.getcwd(),'issue.jpg')
	else:
		print "the screenShot don't have, it may be do the offline issue"
		attachFile=None
	return attachFile
	
if __name__=="__main__":
	try:
		Testing="debug"
		if Testing=="debug":
			statusEmail_debug()
			# screenShotEmail()
			#schedule.every(1).minutes.do(statusEmail_debug)
		else:
			#Must ensure the formor process is killed, otherwise outlook couldn't deliver the mail successfully
			schedule.every().day.at("9:15").do(statusEmail)
			schedule.every().day.at("17:50").do(appExistEmail)
			while True:
				schedule.run_pending()
	except Exception,e:
		traceback.print_exc()
	if raw_input("please press enter key to terminate")=="":
		sys.exit(0)

