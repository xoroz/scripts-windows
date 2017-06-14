'Script to backup PC Desktops 
'Felipe Ferreira
'05/06/2017
'Version 1.0 

'DONE:
'OK - Set List to be backuped in array and text file 
'OK - PST using shadowspawn and robocopy  
'OK - Logfile (list files backuped, time index, time total, size index )
'OK - Clean Robocopy logfile

'OK - Could have txt file with last list (if same day reuse it)
'OK - 1-Create list of files backuped 
'OK - 2-Logs and list on the target share 
'OK - 3-Before doing copy check if already on the today list if it is do not copy it (MUST for PST)
'OK - Permission on Destination folder should be set for User and Admin ONLY:  CACLS files /e /p {USERNAME}:{PERMISSION}
'OK - Limiti bandwidth for .PST files, over 1GB - can be done with robocopy (10m)
'OK - Control max number of backups to run at the same time 

'POSSIBLE BUG:
'If user cancel backup with CONTROL-C lock file will remain and no other backups will run ( maybe ignore if lock if file is from same username)

'TODO: (if we will really use the script)
'Send an email as HTML (2hr)
'Limiti bandwidth only on large files, over 1GB - can be done with robocopy (10m)
'Simple HTML page for status and recent backups  (90m)

'BEST WAY TO RUN IT ?
'(not possible) 1- RemoteExec via psexec and .bat file  '-1. BAT should copy all req files and call vbs script (must run with the logged user profile)
'2- Via GPO, on logon 
'3(not possible)- Configure Remote Task Schedulle on each PC 
'4- Manually, ask user to click on desktop .bat file 
'5-Via GPO, bat to configure schedulle task ? 

'NOTES:
'-No delta backup / no incremental (but robocopy only copy files changed)
'-No encryption (but only the user has access to its folder)
'-No compression
'-User restore (directly via SHARE)
'+Free
'+Easy and simple to Run 
'+Fast 
'+Can backup opened pst files using shadowspwan 
'+Robocopy will only copied files that have been changed 
'-Robocopy will also copy PST everytime (since outlook changes it) (Implemented a check to only copy once a day)

'DOUBTS:
'-PST ONCE a Month OK ? 
'-Maybe ZIP -9 it before sending it and have server process to unzip and fix permissions 

'MAINTANCE(HOWTO)
'Would run remotely from a Task Scheduller to each PC 


'Declare constant Variables 
Const ForReading = 1 : Const ForWriting = 2 : Const ForAppending = 8
		
'Declare Generic Variables 
Dim objfso,ldr,objFolder,objShell,debug,strSearchD,strLog,strLogRobo,strFileList,strTargetG,strUserDir,strUserName,StrFileLock,strLockDir,intConcBkp

'Declare Arrays
Dim arrFiles,arrFileInclude,arrFolderInclude,arrFileLines1

'Declare Counter,Timer
Dim timerS,timerE,timerD,timeDS,timerDT,intCountSize,intCountFiles,intCountFilesCopied

'Initialize Variables 
intCountFiles = 0
intCountFilesCopied = 0
intCountSize = 0
timerS = Timer()
arrFiles = Array()
arrFileLines1 = Array()

Dim t : t = Now : timeStamp =  Right("0" & day(t),2)
timeStampFull =  Right("0" & day(t),2) & "-" & Right("0" & month(t),2) & "-" & Right("0" & year(t),4) & " " & Right("0" & hour(t),2) & ":" &  Right("0" & minute(t),2)

Set objfso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
'gets USERPROFILE
strSearchD=Replace(Replace(objShell.ExpandEnvironmentStrings("%USERPROFILE%"),objShell.ExpandEnvironmentStrings("%USERNAME%"),""), "\\", "")
ldr = objfso.GetFolder(strSearchD)
strSearchD=right(strSearchD,len(strSearchD)-3)
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")

'------------------ EDIT HERE -------------------------
verbose = 1
logging = 1 
strTargetG = "\\%USERDOMAIN%\backup\" 'IMPORTANT: EDIT ALSO LINE 200 INSIDE THE SUB DOBACKUP (MUST END WITH "\")
arrFileInclude = Array("pst","pdf","doc","docx","txt","csv","ods","xlsx","xls","xlsm","ppt","pptx","msg")
arrFolderInclude  = Array("Documents","Desktop")
intConcBkp = 1 ' HOW MANY BACKUPS CAN RUN AT THE SAME TIME 
'------------------ DONE EDIT -------------------------

strLockDir = strTargetG & "CONTROL"
strUserDir = strTargetG & strSearchD 


'START LOCK CONTROLS (IF ANY BACKUP IS CURRENTLY RUNNING QUIT)
Call checkdir(strlockdir)
StrFileLock = strlockdir & "\bkp_" & strusername & "_" & timeStamp & ".lock"

if checkiflock() = 1 then    
 wscript.quit 2
else 
 call createLock()
end if 
'END LOCK CONTROL 


Call checkdir(struserdir)
strLogRobo = strUserDir & "\bkp_robocopy_" & timeStamp & ".log"
strLog = strUserDir & "\bkp_" & timeStamp & ".log"
strFileList = strUserDir & "\bkp_copied_" & timeStamp & ".log"

pt "DEBUG " & vbcrlf & "RoboLog: " & strLogRobo & vbcrlf & "LOG: " & strLog & vbcrlf & " FileList: " & strFileList & vbcrlf & "USENAME: " & strUserName & vbcrlf & "USERDIR: " & strUserDir 

pt vbcrlf & vbcrlf & "---------------------------------------------------------"
pt timeStampFull & vbcrlf 
pt "-Building index for : " & ldr 
pt "--Folder Filters: " 
for each item in arrFolderInclude
 Wscript.StdOut.Write item & " "
next 
pt vbcrlf & "--File Filters: "
for each item in arrFileInclude
 Wscript.StdOut.Write item & " "
next 
pt vbcrlf & "Please Wait..."

'------------MAIN CALLS----------------
Call GetList(ldr)

TimerE = Timer()
TimerD = FormatNumber(TimerE - TimerS,1)
pt vbcrlf & "Found: " &  intCountFiles & " files to backup."
pt "Found: " &  intCountSize & " MBs  to backup."
pt "Index Time: " &  TimerD & " seconds"

timerS = Timer() 

Call Dobackup

TimerE = Timer()
TimerDS = FormatNumber(TimerE - TimerS, 1)
pt vbcrlf &  "Sync/Copy Time: " &  TimerDS & " seconds"
'Call SetACL() - Too much network usage for this cmd (should run local on the server side)
pt "Backup data and logs at: " & strTargetG & strSearchD 
'TimerDT = TimerDS + TimerD
'pt vbcrlf & "Total Time: " &  TimerDT & " seconds"
call RemoveLock()
wscript.quit 0 
'------------END MAIN CALLS----------------



'---------------------------------------------------------------------------------------------------------------------
Sub GetList(fFolder)
on error resume next 
	Dim colFiles,objFile,Subfolder,intLastMod
	Set objfso = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(fFolder)   
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		If objfso.FileExists(objFile) Then    
			'pt objFile.Name
			if FilterFile(objFile.name) = 1 then 
			intSize = formatNumber(objFile.size / 1024 / 1024, 3)
			intLastMod = CDATE(objfile.DateLastModified)						
			'pt "FOUND: " & objFile.Path & " ; " & intSize & " ; " & intLastMod 		
			Wscript.StdOut.Write "." 
			Redim preserve arrFiles(UBound(arrFiles) + 1)
            arrFiles(intCountFiles) = objFile.Path
			intCountFiles = intCountFiles + 1
			intCountSize  = intCountSize + intSize 
			end if 
		End If        
	Next

    For Each Subfolder in objFolder.SubFolders
	'    pt SubFolder
		if FilterFolder(Subfolder) = 1 then 		
			GetList(Subfolder)		
		end if 
    Next
End Sub

Function FilterFile(oFile)	
    'pt "Checking: " & oFile 
	For each ext in arrFileInclude		
		If UCase(objFSO.GetExtensionName(oFile)) = Ucase(ext) Then
		 FilterFile = "1"		
		end if 
	Next  
end Function  

Function FilterFolder(oFolder)	
    Dim intLast    
	For each folder in arrFolderInclude		
		'pt "Folder Filter: " & UCASE(strFolder) & " = " &  UCASE(Folder)
		If instr(UCase(OFolder),Ucase(folder)) Then
	     'pt "Found Folder Filter: " & UCASE(oFolder)
		 FilterFolder = "1"		
		end if 
	Next  
end Function 



Sub Dobackup()
'get array of files to backup and execute robocopy
'get array of pst files and execute hobocopy
	Dim sFiles, item, i, oshell,strDir
	strDir=""
	Set oShell = WScript.CreateObject ("WScript.Shell")
	intTotal = Ubound(arrFiles) 
	pt vbcrlf & "Starting sync/copy " & intTotal + 1 & " files to " & strUserDir 
	For i = 0 to intTotal 
		strTarget = strTargetG
	    strDir=""
	    intLast=0
	    'pt "COPYING ITEM: " & arrFiles(i) & vbnewline		
	    'Separate Filename and Directory , and create target dir path
		arrSplit = Split(arrFiles(i), "\")
		intLast = Ubound(arrSplit)		
		for it = 0  to intLast - 1
			if (it = intLast - 1) then 		
				strdir = strdir  & cstr(arrSplit(it))
				strTarget = strTarget & cstr(arrSplit(it))
			else
				strdir = strdir  & cstr(arrSplit(it)) & "\"
				if (it <> 0) then 						
					strTarget = strTarget & cstr(arrSplit(it)) & "\"
				end if 
			end if 
		next 
		strFilename = arrSplit(intLast)
		Set oShell = WScript.CreateObject ("WScript.Shell")
		if UCASE(Mid(strFilename,len(strFilename)-3,4)) = ".PST" then	 'User ShadowSpawn to open PST file then robocopy it 
			strCmd = "cmd /c shadowspawn.exe " & chr(34) & strdir & chr(34) & " W:\ robocopy /NP /NDL /NJH /NJS /ZB /IPG:300 /R:3 /W:5 /LOG+:" & strLogRobo & " /TEE  W:\ " & chr(34) & strTarget & chr(34) & " " & chr(34) & strfilename & chr(34) 
		else		   
			strCmd = "cmd /c robocopy /NJH /NP /NDL /NJS /ZB /R:3 /W:5 /LOG+:" & strLogRobo & " /TEE " & chr(34) & strdir & chr(34) & " "& chr(34)  & strTarget & chr(34)  & " " & chr(34) & strfilename & chr(34) 
		end if 
		
		if CheckList(strFilename) = 0 then 
			Return = oshell.Run(strCmd,0, true)		
			if ( Return = 1 ) or ( Return = 32769 ) then 			
				pt vbcrlf & "File copied: " & strFilename  				
				ptList strFilename  
				intCountFilesCopied = intCountFilesCopied + 1
				if intCountFilesCopied mod 100 = 0 then
					pt "COPIED: " & intCountFilesCopied & " of  " & intCountFiles & " dataset files. Total: " & intCountSize & " MBs"
				end if
				
			elseif ( return <> 0 ) then 
				pt "ERROR_CODE: " & return 
			end if       
		'elseif CheckList(strFilename) = 1 then 
			'pt "File: " & strFileName & " already backuped today"
			
		end if 
		
	Next 
	
end sub 

function CheckList(strFileCp)
'OPEN strFileList AND IN CASE FILE IS IN THERE (IT WAS BACKUPED TODAY) DO NOT BACKUP 
	Dim objFSOList,objFileList,l,itemL
	l = 0	
	
	Set objFSOList = CreateObject("Scripting.FileSystemObject")		
	if Ubound(arrFileLines1) < 1 then 
		if objfsoList.FileExists(strFileList)   Then
			Set objFileList = objFSOList.OpenTextFile(strFileList, ForReading)
			Do Until objFileList.AtEndOfStream		
				Redim Preserve arrFileLines1(l)
				arrFileLines1(l) = objFileList.ReadLine
				'pt "READLINE: " & objFileList.ReadLine
				l = l + 1
			Loop		
			objFileList.Close
		end if 
	'else		
		'pt "Array Already populated " & Ubound(arrFileLines1) & " itens found!  L: " & l 
	 end if  
	'CHECK IF THE FILE IS IN THE ARRAY/FILELIST 
	for each itemL in arrFileLines1 
		if itemL = strFileCp then			
	'pt strFileCp & "File is in the list"
			CheckList = 1
			Exit Function 
		end if 
	next
	'pt strFileCp & "File is NOT in the list"
	CheckList = 0		
	Wscript.StdOut.Write "." 
end function



sub SetACL()
'Should set ACL permission only for the user to see its folder, runs only if some files were copied to avoid running each time 
'ATTENTION: Very Network Intesnsive (better run central on the server side )
	if intCountFilesCopied > 20 then 
		Set oShell = WScript.CreateObject ("WScript.Shell")	
		pt "Configuring Permssions only for " & strUserName & " on " & strUserDir 
		
		strCmd = "cmd /c icacls " & strUserDir & " /Q /t /c /inheritance:d /grant:r dm\" & strUserName & ":F"	
		Return = oshell.Run(strCmd,0, true)		
		'pt Return 	
		strCmd = "cmd /c icacls " & strUserDir & " /Q /t /c /setowner dm\" & strUserName 
		Return = oshell.Run(strCmd,0, true)		
		'pt Return 
		strCmd = "cmd /c icacls " & strUserDir & " /Q /t /c /remove:g " & chr(34) & "Authenticated Users" & chr(34) & " /remove:g Users /remove:g Everyone"	
		Return = oshell.Run(strCmd,0, true)		
		'pt Return 
	end if 
end sub 

sub checkdir(strD)
    Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT (FSO.FolderExists(strD)) Then
		pt "Creating " & strD	
		Set objShell = CreateObject("Wscript.Shell")
		objShell.Run "cmd /c mkdir " & strD
		WScript.Sleep 500
	end if 
	Set fso = nothing
	Set objShell = nothing 	
end sub 

function checkiflock()
'CHECK IF ANY BACKUP IS CURRENTLY RUNNING, IF SO CANCEL THIS ONE 
	Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT (ObjFSO.FolderExists(strLockDir)) Then
		pt "ERROR - Foler " & strLockDir & " not found!" 
		wscript.quit 2
	else
		Set folder = fso.GetFolder(strLockDir)
		if folder.files.Count <= intConcBkp then
			pt "OK - " & folder.files.Count & " backups running in " & strLockDir
			checkiflock = 0
		else	
			pt "OPS - already " & folder.files.Count & " backups running, Max is " & intConcBkp & ". Maybe Delete Lock file in " & strLockDir
			checkiflock = 1
		end if
	end if 
	Set fso = nothing
end function 

sub createLock()
	Set fso = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFileLock, ForAppending, True )
	pt "LOCK CREATED at " & strFileLock
	fso.close 
	Set fso = nothing 
end sub 

sub RemoveLock()
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile strFileLock
	pt "LOCK REMOVED at " & strFileLock	
	Set fso = nothing 
end sub 

sub pt(txtmsg)
	on error resume next 
	if verbose = 1 then
		wscript.echo txtmsg
	end if
	if logging = 1 then
		Set objFSOL = CreateObject("Scripting.FileSystemObject")
		Set objFileLog = objFSOL.OpenTextFile(strLog, ForAppending, True )
			objFileLog.Writeline txtmsg 
			objFileLog.close 
	end if 
end sub 

sub ptList(txtmsg)	
	Set objFSOL = CreateObject("Scripting.FileSystemObject")
	Set objFileLog = objFSOL.OpenTextFile(strFileList, ForAppending, True )
	objFileLog.Writeline txtmsg 	
	objFileLog.close 
end sub 

