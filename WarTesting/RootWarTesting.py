from selenium import webdriver
import time
import os
import re
import sys
import traceback
import logging
import time
import platform
import schedule

class RootWarTesting(object):
	def __init__(self):
		self.properties=Property("config.ini").parse()
		self.ipAddress=self.properties['ipAddress']
		self.adminPwd=self.properties['adminPwd']
		self.clientPath=self.properties['clientPath']
		self.serverPath=self.properties['serverPath']
		self.executable_path=self.properties['executable_path']
		self.Version=self.properties['Version']
		self.Version=6.0
		self.CSCFTPServerPath=self.properties['CSCFTPServerPath']
		self.Build=self.properties['Build']
		self.addCommand=os.path.join(self.clientPath,"bin","cscadd")
		self.regCommand=os.path.join(self.clientPath,"bin","cscreg")
		self.content=self.getContent("courier.cfg")
		self.platF=platform.platform()
		self.initValue=self.getInitIDValue()+1
		self.courierName="test"+str(self.initValue)
		self.currentDate=time.strftime("%Y%m%d",time.localtime())
		self.currentPath=os.getcwd()
		self.fp=open("Result.txt","a")
		if self.platF.startswith("Win"):
			os.system("type nul>%s"%os.path.join(self.currentPath,"log","log%s.txt"%self.currentDate))
			self.KillPid("java.exe")
			self.KillPid("tomcat9.exe")
		else:
			os.system("touch %s"%os.path.join(self.currentPath,"log","log%s.txt"%self.currentDate))
		logger = logging.getLogger(__name__)
		logger.setLevel(level = logging.INFO)
		self.logger=logger

		
	def judgelatestRootWarExists(self):
		if not os.path.exists(os.path.join(self.CSCFTPServerPath,self.Build,self.currentDate)):
			self.printLog("The latest root.war hasn't deployed")
			sys.exit(0)
			
			
	def getURL(self):
		self.driver=webdriver.Ie(executable_path=self.executable_path)
		self.driver.get(self.ipAddress)
		self.driver.maximize_window() 
		
	def getInitIDValue(self):
		dirList=os.listdir(os.path.join(self.clientPath,"couriers"))
		newDirList=[]
		for i in dirList:
			if re.match(r"test(\d+)",i):
				newDirList.append(int(re.match(r"test(\d+)",i).group(1)))
		if newDirList!=[]:
			maxValue=max(newDirList)+1
		else:
			maxValue=0
		return maxValue
		
	def getContent(self,file):
		fp=open(file,"r")
		content=fp.readlines()
		fp.close()
		return content
		
	def checkHeartBeat(self):
		os.system("%s %s"%(os.path.join(self.clientPath,"bin","cscstart"),self.courierName))
		logPath=os.path.join(self.clientPath,"couriers",self.courierName,"logs","%s_inbound.log"%self.courierName)
		count=0
		while True:
			if os.path.exists(logPath):
				fp=open(logPath,"r")
				break
			else:
				time.sleep(15)
				count+=1
			if count==10:
				if os.path.exists(logPath):
					fp.close()
				raise Exception("The inbound file couldn't be generated after courier is started up") 
		count=0
		time.sleep(240)
		while True:
			if len(re.findall(r"heart beat Msg",fp.read()))>=2:
				self.printLog("The heartbeat msg is checked successfully")
				fp.close()
				break
			else:
				fp.seek(0,0)
				fp.flush()
				time.sleep(120)
				count+=1
			if count==4:
				fp.close()
				raise Exception("The heart beat message hasn't generated in the inbound file") 
				
			
	
	def createCourier(self):
		self.printLog("Start to do the action for the courier %s\n"%self.courierName)
		if self.platF.startswith("Win"):
			os.system("type nul>temp.txt")
		else:
			os.system("touch temp.txt")
		if os.system("%s %s 2"%(self.addCommand,self.courierName))==0:
			self.logSteps("courier %s is added"%self.courierName)
		else:
			raise Exception("It creates the courier %s failed"%self.courierName)
		
		with open(os.path.join(self.clientPath,"couriers",self.courierName,"courier.cfg"),"w") as fp:
			fp.write("".join(self.content).replace("test0",self.courierName))
		if os.system("%s %s>>temp.txt"%(self.regCommand,self.courierName))!=0:
			raise Exception("The client complete the courier %s register failed"%self.courierName)
		temp=self.getContent("temp.txt")
		for item in temp:
			self.logSteps(item)
		self.printLog("Finish to do the action for the courier %s\n"%self.courierName)
		os.system("rm -rf temp.txt")
		
	def login(self):
		self.printLog("Start to login the website")
		#the below is for win2016 web security validation, 
		#if use different platform and different browser, the below code may be need to be updated.
		try:
			js='document.getElementById("overridelink").click()'
			self.driver.execute_script(js)
		except Exception, e:
			pass
		time.sleep(4)
		try:
			username=self.driver.find_element_by_name("username")
		except Exception,e:
			raise Exception("The login page couldn't be found")
			
		username.clear()
		username.send_keys("administrator")
		password=self.driver.find_element_by_name("password")
		password.send_keys(self.adminPwd)
		js='document.getElementsByClassName("btn-primary hide-focus")[0].click()'
		self.driver.execute_script(js)
		time.sleep(4)
		self.printLog("Logined the website")
		
	def checkElementExists(self,elementXpath):
		try:
			self.driver.find_element_by_xpath(elementXpath)
			return True
		except Exception,e:
			return False
		
	def textCompare(self,elementXpath,expected):
		textValue=self.driver.find_element_by_xpath(elementXpath).text
		textValue=textValue.encode("utf-8").strip()
		if textValue==expected:
			return True
		else:
			raise Exception("The value isn't match,the actual value is %s,the expected value is %s"%(textValue,expected))
			return False

	def KillPid(self,app):
		if self.platF.startswith("Win"):
			os.system("tasklist>taskList.txt")
			List=[]
			with open("taskList.txt","r") as fp:
				for line in fp:
					if app in line and re.findall(r"\s(\d+)\s",line,re.M)[0] not in List:
						List.append(re.findall(r"\s(\d+)\s",line,re.M)[0])
			for item in List:
				os.system("taskkill/F /pid %s" %item)
	
	def approveCour(self,courierName):
		#switch to the new courier
		self.printLog("Start to approve the courier")
		#expand the menu
		js='document.getElementsByClassName("btn-icon application-menu-trigger hide-focus")[0].click()'
		self.driver.execute_script(js)
		#expand the courier section
		js='document.getElementsByClassName("btn hide-focus")[0].click()'
		self.driver.execute_script(js)
		time.sleep(4)
		#click the new courier
		js='document.getElementsByClassName("accordion-header list-item hide-focus")[0].click()'
		self.driver.execute_script(js)
		time.sleep(4)
		result=self.checkElementExists("//span[@class='page-title']")
		if result:
			self.textCompare("//span[@class='page-title']","New Couriers")
		else:
			raise Exception("The New Courier title couldn't be seen,it may be switch page failed")
		
		#select the first record
		js='document.getElementsByClassName("datagrid-checkbox datagrid-selection-checkbox")[0].click()'
		self.driver.execute_script(js)
		#apprvoe part
		js='document.getElementById("delete-item").click()'
		self.driver.execute_script(js)
		time.sleep(4)
		SystemId=self.driver.find_element_by_name("acd_systemId")
		SystemId.send_keys(courierName)
		KeyPass=self.driver.find_element_by_name("acd_password")
		KeyPass.send_keys("P@ssword01")
		KeyConfirmPass=self.driver.find_element_by_name("acd_confirmPassword")
		KeyConfirmPass.send_keys("P@ssword01")
		js='document.getElementsByClassName("btn-modal-primary hide-focus")[0].click()'
		self.driver.execute_script(js)
		for i in range(10):
			time.sleep(4)
			try:
				self.driver.find_element_by_id("message-title")
				break
			except Exception,e:
				pass
		else:
			raise Exception("The approve loading failed")
			
		js='document.getElementById("modal-button0").click()'
		self.driver.execute_script(js)
		time.sleep(4)
		try:
			self.driver.find_element_by_id("message-title")
			raise Exception("The confirm alert hasn't been closed")
		except Exception,e:
			pass
		try:
			firstCour=self.driver.find_element_by_xpath("//table[@role='grid']/tbody/tr[@aria-rowindex='1']/td[2]")
			raise Exception("The approved courier hasn't been removed from the new courier page")
		except Exception,e:
			pass
		self.printLog("Approved the courier")
		
		
	def registerCourierValueJudge(self,courierName):
		self.printLog("Start to judge whether the courier is created successfully")
		js='document.getElementsByClassName("accordion-header list-item hide-focus")[1].click()'
		self.driver.execute_script(js)
		time.sleep(4)
		result=self.checkElementExists("//span[@class='page-title']")
		if result:
			self.textCompare("//span[@class='page-title']","Registered Couriers")
		else:
			raise Exception("The Registered Courier title couldn't be seen,it may be switch page failed")
		try:
			firstCour=self.driver.find_element_by_xpath("//table[@role='grid']/tbody/tr[@aria-rowindex='1']/td[2]")
		except Exception,e:
			raise Exception("The approved failed, it couldn't find that courier %s in the register page"%self.courierName)
		self.textCompare("//table[@role='grid']/tbody/tr[@aria-rowindex='1']/td[2]",courierName)
		self.printLog("The courier is approved successfully")
		
		
	def closeDriver(self):
		self.driver.close()
		
		
	def cleanRegCourier(self,courierName):
		self.printLog("Start to clean up the environment")
		os.system("rm -rf %s"%os.path.join(self.clientPath,"couriers",courierName))
		fp=open(os.path.join(self.clientPath,"couriers","Start.cfg"),"w")
		fp.write("")
		fp.close()
		path=os.getcwd()
		os.chdir(os.path.join(self.serverPath,"lws","new"))
		os.system("rm -rf *")
		os.chdir(path)
		os.system("rm -rf %s"%os.path.join(self.serverPath,"lws","registered","default",courierName))
		os.system("rm -rf %s"%os.path.join(self.serverPath,"lws","revoked",courierName))
		os.system("rm -rf %s"%os.path.join(self.serverPath,"certs","%s.crt"%courierName))
		os.system("rm -rf %s"%os.path.join(self.serverPath,"certs","%s.csr"%courierName))
		os.system("rm -rf %s"%os.path.join(self.serverPath,"certs","%s.jks"%courierName))
		os.system("rm -rf %s"%os.path.join(self.serverPath,"certs","%s.log"%courierName))
		fp=open(os.path.join(self.serverPath,"lws","registered","courierversion.cfg"),"w")
		fp.write("")
		fp.close()
		fp=open(os.path.join(self.serverPath,"lws",".tmp_sysid.ctr"),"w")
		fp.write("10000\n")
		fp.close()
		self.printLog("Finish the clean up")
		
	def rootWarSetting(self):
		self.printLog("Start to rootWar setting")
		os.system('net stop "Infor Cloverleaf(R) Secure Courier %s Server"'%self.Version)
		RootWarSourcePath=self.CSCFTPServerPath%self.currentDate
		RootWarTargetPath=os.path.join(self.serverPath,"tomcat","regapps","ROOT.war")
		RootTargetPath=os.path.join(self.serverPath,"tomcat","regapps","ROOT")
		ESAPIPath=os.path.join(RootTargetPath,"WEB-INF","classes","ESAPI.properties")
		os.system("rm -rf %s"%RootWarTargetPath)
		os.system("rm -rf %s"%RootTargetPath)
		os.chdir(os.path.join(self.serverPath,"tomcat","regapps"))
		result=os.system("wget %s"%RootWarSourcePath)
		os.chdir(self.currentPath)
		if result==1:
			raise Exception("The date %s root.war hasn't deployed"%self.currentDate)
		os.system('net start "Infor Cloverleaf(R) Secure Courier %s Server"'%self.Version)
		while True:
			if os.path.exists(RootTargetPath):
				while True:
					if os.path.exists(ESAPIPath):
						#as ESAPI.properties content generation finished need some time
						time.sleep(4)
						break
				content=self.getContent(ESAPIPath)
				with open(ESAPIPath,"w") as fp:
					fp.write("".join(content).replace("ROOTPATH",self.serverPath))
				break
				
		self.printLog("Finished the rootWar setting")
		return result
	def printLog(self,content):
		print content
		self.logSteps(content)
		
	def logSteps(self,content):
		'''judge whether need to print in the console, and logger the records in the log file'''
		if len(self.logger.handlers)==0:
			handler = logging.FileHandler(os.path.join(self.currentPath,"log","log%s.txt"%self.currentDate))
			self.logger.addHandler(handler)
		if content.startswith('\n'):
			self.logger.info('\n'+time.strftime("[%m/%d/%Y %H:%M:%S]",time.localtime())+" "+content.strip())
		else:
			self.logger.info(time.strftime("[%m/%d/%Y %H:%M:%S]",time.localtime())+" "+content.strip())
	
def mainCheck():
	rwt=RootWarTesting()
	try:
		rwt.rootWarSetting()
		rwt.createCourier()
		rwt.getURL()
		rwt.login()
		rwt.approveCour(rwt.courierName)
		rwt.registerCourierValueJudge(rwt.courierName)
		rwt.checkHeartBeat()
		os.system("%s %s"%(os.path.join(rwt.clientPath,"bin","cscstop"),rwt.courierName))
		rwt.fp.write("%s log passed\n"%rwt.currentDate)
		rwt.fp.flush()
	except Exception,e:
		#print the exception
		traceback.print_exc()
		#log the exception
		rwt.logger.exception("Exception Logged")
		rwt.fp.write("%s log failed\n"%rwt.currentDate)
		rwt.fp.flush()
		os.system("%s %s"%(os.path.join(rwt.clientPath,"bin","cscstop"),rwt.courierName))
	finally:
		try:
			rwt.closeDriver()
		except Exception,e:
			traceback.print_exc()
			rwt.logger.exception("Exception Logged")
			rwt.fp.write("%s log failed\n"%rwt.currentDate)
			rwt.fp.flush()
		rwt.fp.close()
		rwt.logger.removeHandler(logging.FileHandler(os.path.join(rwt.currentPath,"log","log%s.txt"%rwt.currentDate)))
		rwt.cleanRegCourier(rwt.courierName)
		

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
		
def cleanSpecificCourier(courierName):
	rwt=RootWarTesting()
	rwt.cleanRegCourier(courierName)
	
	
if __name__ == "__main__":
	Testing="formal"
	
	if Testing=="debug":
		# schedule.every().day.at("18:18").do(rwt.mainCheck)
		# schedule.every(1).minutes.do(rwt.mainCheck)
		# cleanSpecificCourier("test1")
		mainCheck()
	else:
		schedule.every().day.at("19:20").do(mainCheck)
		while True:
			schedule.run_pending()

