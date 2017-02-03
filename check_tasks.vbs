'
' Script to get failed schedulle jobs (plugin for Nagios/Centreon)
' by Felipe Ferreira
'REQUIRES: schtasks.exe
'Tested only Windows 2003

'for configuring NSCP/NRPE create a .bat file because when called like cscript script/file.vbs does not work!

Option Explicit 
Dim objFSO, objTextFile, strNextLine, arrServiceList
Dim objshell, sCommand, strTmpFIle, intCount
Dim intSize, intErr, intStatus, strName, strmsg, blnLastRun,strCurDir 

Const ForReading = 1 
Const EVENT_ERROR = 1 
intCount = 0
intSize = 0 
intErr = 0

Set objshell = Wscript.Createobject("WScript.Shell") 
strCurDir    = objshell.CurrentDirectory
strTmpFile = strCurDir & "\schedtaskmon.txt"

sCommand = "cmd /C schtasks.exe /query /FO CSV /V > " & strTmpFile  
objshell.Run sCommand 
WScript.Sleep 1000 

Set objFSO = CreateObject("Scripting.FileSystemObject") 
 
If objFSO.FileExists(strTmpFile) Then 
  Set objTextFile = objFSO.OpenTextFile(strTmpFIle, ForReading) 
  objTextFile.SkipLine 
  objTextFile.SkipLine 
else
 wscript.echo "UNKOWN - File "   & strtmpfile & " could not be found"
 wscript.quit 3
End If

Do Until objTextFile.AtEndOfStream
  intCount = 0  
  strNextLine = objTextFile.ReadLine   
  arrServiceList = Split(strNextLine,",")
  intStatus = len(arrServiceList(3))
  if not (intStatus = 2 ) then    
   if ( len(arrServiceList(7)) = 3 ) then 
    strName = arrServiceList(1)  
    blnLastRun = arrServiceList(7)
   end if 
  end if    
  'wscript.echo " NAME: " &  strName & " Last Run: "     & blnLastRun 
  if  instr(blnLastRun,"1" ) then 
   'wscript.echo " NAME: " &  strName & " Last Run: "     & blnLastRun 
   'strMsg = " NAME: " &  strName & " Last Run: "     & blnLastRun & strmsg
   strMsg =  strName & " " & strmsg
   intErr = intErr + 1
  end if    
Loop


if ( intErr > 0 )then
 wscript.echo "CRITICAL - Found " & intErr & " schedulle task job(s) that failed " & strmsg & "|job_error=" & intErr 
 wscript.quit 2
else
 wscript.echo "OK - No jobs fail found |job_error=" & intErr 
 wscript.quit 0
 end if 
