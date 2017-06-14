'Script to check if a specific patch is installed
'by Felipe Ferreira 06/2017


Const ForReading = 1 : Const ForWriting = 2 : Const ForAppending = 8
Dim verbose,logging,blnFound
Dim t : t = Now : timeStamp =  Right("0" & day(t),2)
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

'--------EDIT HERE----------
verbose = 1
logging = 1 
strLogDir="\\%USERDOMAIN%\netlogon\"
aComputers=Array("bobo-nb", "bobao-nb", "bobinha-nb", "bobapc-nb")
PATCH1="KB4012215"
PATCH2="KB4012212"
'--------DONE EDIT----------

timeStampFull =  Right("0" & day(t),2) & "-" & Right("0" & month(t),2) & "-" & Right("0" & year(t),4) & " " & Right("0" & hour(t),2) & ":" &  Right("0" & minute(t),2)
strLog = strLogDir & "check_patch_" & timestamp & ".log"

pt  vbcrlf & "------------------------------------------------------------------"
pt  timeStampFull
pt "Checking for Patches: " & PATCH1 & " and " & PATCH2 

for each strComputer in aComputers 
	'pt  vbcrlf & "Checking: " & strComputer & vbcrlf
	if ping(strComputer) = True  then 
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering",,48)
		If Err.Number <> 0 Then
			pt "WMI Connection error #" & Err.Number & " on "& strComputer & " : " & Err.Description		
			Err.Clear
			exit for 
		end if 
		
		For Each objItem in colItems
		 'pt strComputer & " - " & objItem.HotfixID 
		 wscript.StdOut.Write "." 
		 If ((objItem.HotfixID = PATCH1) or (objItem.HotfixID = PATCH2)) then		  
		  pt  "OK - " & strComputer & " patch: " & objItem.HotfixID & " found."
		  blnFound = 1
		  exit for 
		 End If
		next 
	end if 
	if blnFound <> 1 then 
	 pt "CRITICAL - " & strComputer & " does not have the Patch installed"
	end if 
	blnFound = 0 
Next

'------------------SUB AND FUNCTIONS-----------------
Function Ping( myHostName )
    Dim colPingResults, objPingResult, strQuery
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & myHostName & "'"
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )
    For Each objPingResult In colPingResults
        If Not IsObject( objPingResult ) Then
            Ping = False
			pt "ERROR - " & myHostName & " is unreacheable"
        ElseIf objPingResult.StatusCode = 0 Then
            Ping = True
        Else
            Ping = False
			pt "ERROR - " & myHostName & " is unreacheable"
        End If
    Next
    Set colPingResults = Nothing
End Function

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
