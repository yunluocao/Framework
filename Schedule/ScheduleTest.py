#encoding=utf-8
import schedule
import traceback
import sys

class ScheduleTest(object):
	def __init__(self):
		self.properties=Property("config.ini").parse()
		self.timeTest=self.properties['timeTest']



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


def printLog():
	print "abcd"
	
try:
	Testing="formal"
	if Testing=="debug":
		pass
		#schedule.every(1).minutes.do(statusEmail_debug)
	else:
		sch=ScheduleTest()
		orginal=sch.timeTest
		#Must ensure the formor process is killed, otherwise outlook couldn't deliver the mail successfully
		schedule.every().day.at(sch.timeTest).do(printLog)
		while True:
			sch=ScheduleTest()
			if sch.timeTest!=orginal:
				schedule.every().day.at(sch.timeTest).do(printLog)
				orginal=sch.timeTest
			schedule.run_pending()
			# sch=ScheduleTest()
except Exception,e:
	traceback.print_exc()
if raw_input("please press enter key to terminate")=="":
	sys.exit(0)