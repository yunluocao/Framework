Sub Business

1.Grab the risk severity and risk name, and add some information
2.Compare the risk name whether included the summary, otherwise add the step1 in the new line
3.If the risk name existed, go the related sheet, add the new line,and add the risk content into the related sheet,
else create new sheet, and add the description name in the first line first column, and add the html name in the first line second column, and add the url in the first line third column

Html Handle class
function:
(1)grab the risk name list, and risk priority
(2)grab the risk content,return cont
(3)grab the risk information
(4)grab the risk name url
Note:we could use risk priority and risk name as key, the other could be value.
(5)combine the risk headline(risk Name+risk Priority+risk information)

Excel handle class
function:
(1)Create one excel and summary sheet including the title headline.(It should handle the known exception(file exists and file is opened ) as well)
(2)Compare the risk name existed
(3)Compare the risk name whether included the summary, otherwise add the step1 in the new line within sample information.
(4)Modify the exist sheet information if the sheet existed,update the description line's section and URL columns.
(5)Add the new sheet which named as No.value, add the title headline and add the risk description,section and url.
(6)Go the related sheet name,add the description content in the max row





