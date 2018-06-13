This tool is to generate the report automatically for the remediation report

1.It will initial the dictionary at first, including the source html folder, merge folder,backup folder, and saved data folder, and dict files

Note: Both the source folders and the subfolders of the merge folder include each module folders.

When the source folders aren't exist, it needs to run at first to create the module folder, then put the html files manually, and run the tool script again to generate the report.

Though the tool create the standard module folders, it also could create some temporary folders, the step 2 could filter the related folders.


2.According to the merge flag to find the related dictionary folder.

3.If it needs to be mereged, the risk data dict, url dict, hyperlink dict,section dict and description dict 
would get the data from the dict files, and create new dictionary for the merged dict,otherwise it would create new dictionary. Then Parse the html data by BeautifulSoup API to add the related dict.

4.If it needs to be merged, it would load the exist report excel, otherwise it would create new excel

5.Update the summary data and create the related risk sheet and hyperlink

6.Update the related risk sheet data from the related dict data.

7.Format all of the sheet

8.Back up the orginal dict and report files,save the report excel, then merge the merged data to the risk dict data, and finial update all of the dict files

9.According to the moveFlag to move the merged html to the related the source module folder.


This tool needs at least to pass two parameters: "BuildVersion","Phase"
The below flags are optional:
(1)mergeFlag default value is None, or put as "", it would according to the merged folder whether exists any html files, then set the mergeFlag True or False.

False: It would do the merge branch.
True: It would do the initial branch.

If the user sets mergeFlag as False manually, it always do the initial report branch even the merge folder has html files.

(2)moveFlag default value is None, means not to move the html files.
It only move the merged flags to then related source module folders only when the moveFlag as True.
When move the files, if the file has same name as the source file name, it would move the source file to the backup folder at first, then continue to move.

(3)cmdFlag default value is True, means print some information in the cmd console.
It only print some information in the cmd console when the cmdFlag as True, otherwise it won't print some information.
No matter the cmdFlag is True or not, it always save the value in the log.txt
