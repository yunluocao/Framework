This tool is to generate the report automatically for the remediation report
Please pip install bs4 and openpyxl api at first

1.Set the parameters for the class object initial under the __main__

(1)BuildVersion is mandatory, it's for the build version

(2)Phase is mandatory, it's for the build phase

(3)Reporter default is None, it's for the reporter name

(4)DevOwner default is None, it's for the dev owner name

(5)mergeFlag default value is None, or put as "", it would according to the merged folder whether exists any html files, then set the mergeFlag True or False.

True: It would do the merge branch.
False: It would do the initial branch.

Note:
If the user sets mergeFlag as False manually, it always do the initial report branch even the merge folder has html files.

(6)moveFlag default value is None, means not to move the html files.

True: It moves the merged htmls to the related source module folders.
False or default None: It doesn't move the merged htmls.

Note:
When move the files, if the file has same name as the source file name, it would move the source file to the backup folder at first, then continue to move.

(7)cmdFlag default value is True, means print some information in the cmd console.

True: It prints the details information in the cmd console
False:It doesn't print the details information in the cmd console

Note:
No matter the cmdFlag is True or not, it always save the value in the log.txt

Some internal parameters:

(8)CurrentLocation is set for the current report path

(9)ModuleFolderList default have the below module folders:
['CAA','CLAPI','CLRESTWS','SITEDOC','WEBSERVICES','CLWIZARD']


2.It needs to run the tool at first to create the module folders automatically.

3.Put the html folders in the related module folder.

4.Run it to create the report excel(such as:CIS621RC1_OWASP_Scanning_Issue.xlsx)

5.If need to do the merge action, it could put the html files into the related module folder under the /CurrentLocation/Merge/.

6.Run it to create the report excel(such as:CIS621RC1_OWASP_Scanning_Issue.xlsx)

Note: You could find the orginal report file under the backup folder under the /CurrentLocation
