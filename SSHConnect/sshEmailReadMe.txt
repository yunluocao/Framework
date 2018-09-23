#This is for report the daliy testing status and the existance of the execution tool by email

Steps:
1.Config the sshEmailCon.ini
2.Run the sshEmail.exe or run the commond "python sshEmail.py"


The tool includes two part:

(1)daliy basic testing check:
It retrive the yesterday date as the script used the local outlook to retrive the log to avoid the time gap

Use the ssh api to run the remote script ResultFile.py which exists in the remote C:\Users\Administrator,
then get the status and log content

According to the status to set the email body value,if the status is passed, then mentioned pass and give the Root.war ftp address.
If the status is failed, then mentioned failed, and give the log content.

And next use the outlook to mail it, the subject also include the status.
(2)the existance of the execution tool check:
Use the ssh api to run the remote script AppCheck.py which exists in the remote C:\Users\Administrator,
then get the response

According to the response value, if the response include the information, then use the outlook to mail the related person.


(3)Schedule
Make two tasks for the above functions, and schedule it.










