'Check if the logfile is being updated, if not restart the batch file 
'Can be used to monitor any batch file, but should be called from here 
'edit variables: directory, and logfile extension, maxoldtime, batchfile to restart 
'Tested Windows 10
'Felipe Ferreira 01/2018

Option Explicit
Const ForReading = 1 : Const ForWriting = 2 : Const ForAppending = 8

dim FILEDIR,FILEEXT,strLog,strLogCheck
dim starttime,EndTime,timeStamp,dtmDateDiff,maxtime
dim strSMTPTo,strAttachment,strTextBody,verbose,logging,debug,WshShell
dim strComputerName,wshNetwork,blnCheck,blnError,dtmDate
dim strRunCMD, strKillCMD
blnCheck = 0
blnError = 0

'TODO - LOG error and EMAIL 

'---------- EDITARE QUI
FILEDIR="c:\temp\logs\"
FILEEXT="log"
maxtime=5 'if log is not update in this X minutes will kill the process and start a new one 
strSMTPTo = "fel.h2o@gmail.com"
strRunCMD="cmd.exe /k START " & chr(34) & "MINING" & chr(34) & " /ABOVENORMAL call c:\users\ferreira\runtest.bat"            
strKillCMD="taskkill /F /IM cmd.exe /FI " & chr(34) & "WINDOWTITLE eq MINING*" & chr(34) & " /T" 
strLog="c:\temp\logs\checklog.log"

verbose=1 '1 on, 0 off
logging=0 '1 on, 0 off
debug=0  '1 on, 0 off ' will send e-mail each time with the files 
'---------- FINE EDIT

Set wshNetwork = WScript.CreateObject( "WScript.Network" )
strComputerName = wshNetwork.ComputerName

StartTime = Timer()


'------ MAIN ---------
checklog FILEDIR,FILEEXT 
if blnCheck = 1  then 
	pt "Found File: " &  strLogCheck &  " Last Updated in minutes: " & dtmDateDiff & " Last modified at: " & dtmDate 
end if 

if dtmDateDiff > maxtime then 
	pt "Log file " & strLogCheck & " has not been updated in the last  " & maxtime & " minutes "
	pt "RESTARTING PROCESS "
	run strKillCMD
	run strRunCMD
end if 

'E-MAIL IF DEBUG AND CHANGES
if ( debug = 1 ) Then  
 EndTime = Timer()
 pt timeStamp & " - Script took: " & FormatNumber(EndTime - StartTime, 2) & " seconds to finish" 
 strSubject = "DEBUG - LOG FROM - " & strComputerName 
 strAttachment = strLog
 sendmail 
end if 

wscript.quit 0 

'------ END MAIN ---------
function checklog(strdir,strExt) 
	Dim oLstF : Set oLstF = Nothing
	Dim oFile
	Dim goFS  : Set goFS    = CreateObject("Scripting.FileSystemObject")
	For Each oFile In goFS.GetFolder(strdir).Files
      If strExt = LCase(goFS.GetExtensionName(oFile.Name)) Then
         If oLstF Is Nothing Then 
            Set oLstF = oFile ' the first could be the last
         Else
            If oLstF.DateLastModified < oFile.DateLastModified Then
               Set oLstF = oFile		   
            End If
         End If
      End If
	Next
	If oLstF Is Nothing Then
		pt "no " & strExt & " found"
		blnError = 1 
	Else		
		strLogCheck = oLstF.Name
		dtmDate = CDate(oLstF.DateLastModified)
		dtmDateDiff = DateDiff("n", dtmDate, Now)
		blnCheck = 1 	    	
	End If	
end function 


function run(strC)
	Set WshShell = CreateObject("WScript.Shell")	
	'WshShell.Run strC, 1, True	
	WshShell.Exec strC
end function 

	
Function GetStamp()
    Dim t 
    t = Now    
	timeStamp = Year(t)  & "." & _
    Right("0" & Month(t),2)  & _
	Right("0" & Day(t),2)  & "." & _
    Right("0" & Hour(t),2) & _
    Right("0" & Minute(t),2)     
End Function

sub pt(txtmsg)
	on error resume next 
	if verbose = 1 then
		wscript.echo txtmsg
	end if
	'if debug = 1 then
	 strTextBody = strTextBody & " " &  txtmsg & vbCrLf
	'end if 
	if logging = 1 then
		Set objFSOL = CreateObject("Scripting.FileSystemObject")
		Set objFileLog = objFSOL.OpenTextFile(strLog, ForAppending, True )
			objFileLog.Writeline txtmsg 
			objFileLog.close 
	end if 
end sub 

sub pte(txtmsg)
	on error resume next 
	wscript.echo txtmsg
	Set objFSOL = CreateObject("Scripting.FileSystemObject")
	Set objFileLog = objFSOL.OpenTextFile(strLogE, ForAppending, True )
	objFileLog.Writeline txtmsg 
	objFileLog.close 
	strTextBody = strTextBody & " " & txtmsg & vbCrLf
	objFileLog.close 	
end sub

function sendmail() 
on error resume next 
 strSMTPFrom = "checktimbra@datamanagement.it" 
 strSMTPRelay = "mail2.datamanagement.it" 
 Set oMessage = CreateObject("CDO.Message")
 oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
 oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
 oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
 oMessage.Configuration.Fields.Update
 oMessage.Subject = strSubject
 oMessage.From = strSMTPFrom
 oMessage.To = strSMTPTo
 oMessage.TextBody = strTextBody
 oMessage.AddAttachment strAttachment
 oMessage.Send
 wscript.echo "E-mail sent to: " & strSMTPTo
end function 
