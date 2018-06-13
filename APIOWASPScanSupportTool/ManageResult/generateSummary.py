#encoding=utf-8
import os
import json
from openpyxl.workbook import Workbook


List=os.listdir(os.getcwd())
NewList=filter(lambda x:"." not in x,List)

wb = Workbook()
sheet=wb.create_sheet("Summary")
sheet=wb["Sheet"]
wb.remove(sheet)
SumTitleLine=["No","Name","Status","Owner"]
sheet=wb["Summary"]
for index,rowValue in enumerate(["SumTitleLine"]+NewList):
	if index==0:
		sheet.append(SumTitleLine)
	else:
		rowContent=[str(index),rowValue]
		sheet.append(rowContent)
sheet.column_dimensions['B'].width=40
wb.save("OWASP_Scan_Staus.xlsx")
