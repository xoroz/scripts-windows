'list all local users and get their respective groups
'by Felipe Ferreira Jan 2017


on error resume next 
Const Quote = """"

Dim Script, ScriptRun
Dim intLastLogin, strStatus, strDescription
Script    = WScript.FullName
ScriptRun = LCase(Mid(Script, InStrRev(Script, "\") + 1))

If Not Scriptrun = "cscript.exe" Then
	CreateObject( "WScript.Shell" ).Run "cscript //nologo """ & WScript.ScriptFullName
	WScript.Quit
End If

Const Append = 8
Dim objFSO, objFile
Set Fso          = CreateObject("Scripting.FileSystemObject")
Set InputFile    = fso.OpenTextFile("servers.Txt")

strErrorLogFile  = "Errors-" & Year(Date) & Month(Date) & Day(Date) & ".log"
Set objErrorFSO  = CreateObject("Scripting.FileSystemObject")
Set objErrorFile = objErrorFSO.OpenTextFile(strErrorLogFile, Append, True)

strLogFile       = "LocalUsers-" & Year(Date) & Month(Date) & Day(Date) & ".log"
WScript.Echo "[i] Beginning Search"

Do While Not (InputFile.AtEndOfStream)
	On Error Resume Next
	Set objNetwork = Wscript.CreateObject("Wscript.Network")
	strDomain      = objNetwork.ComputerName	
	
	Set objFSO  = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(strLogFile, Append, True)

	strComputer   = InputFile.ReadLine
	WScript.Echo ""
    WScript.Echo "-----------------------------------------------------"
	WScript.Echo "[i] Checking Server: " & strComputer
	' Check if machine is alive/dead.
		strPingStatus = PingStatus(strComputer, strDomain)

	If strPingStatus = "Success" Then
		Wscript.Echo "[i] Successful pinging : " & strComputer
	Else
		objErrorFile.Write strComputer & ": " & strPingStatus & vbCrLf
		Wscript.Echo "[-] Failure pinging " & strComputer & ": " & strPingStatus
		Exit Do
	End If		
    
'LOOP THRU USERS
    WQL="SELECT * FROM Win32_UserAccount"
    Set objLWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\" & strComputer & "\root\cimv2")
    Set colLocalItems = objLWMIService.ExecQuery(WQL)
    For Each objItem in colLocalItems 				
		strGRoup = ""
		intLastLogin = ""			  
	    strUser = objItem.Name				
		strDescription = objItem.Description
		strStatus = objItem.Status		
		'intID = objItem.SID
		WScript.Echo "-----------------------------------------------------"
		WScript.Echo "[i] Users found: " & strUser			
' Win 2003,2000 only 		
		Set colGroups = GetObject("WinNT://" & strComputer & "") 
		colGroups.Filter = Array("group") 
		For Each objGroup In colGroups 
			'Wscript.Echo objGroup.Name  
			For Each objUser in objGroup.Members 
			 if (isObject(ObjUser) and instr(objUser.Name,strUser)) then 
				'Wscript.Echo vbTab & objUser.Name 
				if ( strGroup = "" ) then 
					strGroup = objGroup.Name 
				else
					strGroup = strGroup & ", " & objGroup.Name 
				end if 
				
			 end if 
			Next 
		Next 
		WScript.Echo "[i] Users Group found: " & strGroup			  
		
' should only get locallastlogin for Local users!			
		if instr(objItem.LocalAccount, True) then
		on error resume next 
		Set objUserLocal = GetObject("WinNT://" & strComputer & "/" & strUser)			 
			 If Err = 0 Then
			    intLastLogin = objUserLocal.LastLogin			 
		 		WScript.Echo "[i] Local Last Login: " & intLastLogin				
			 Else			   
			    intLastLogin = ""
				WScript.Echo "[i] X Local Last Login: " & intLastLogin
			 End If
			 On Error GoTo 0	
'now should get Local User status and Description using WMI 
		 WScript.Echo "[i] Local User Status: " & strStatus
		 WScript.Echo "[i] Local User Description: " & strDescription
		else			 
		 intLastLogin = ""			  
		end if		
			objFile.Write strComputer & ";" & strUser & ";" & strGroup & ";" & strStatus & ";" & intLastLogin & ";" & strDescription & vbCrLf					
		WScript.Echo "[i] Done checking User: " & strUser		
	Next

	WScript.Echo "-----------------------------------------------------"
	WScript.Echo "Finished searching: " & strComputer
	objFile.Close
Loop

WScript.Echo "-----------------------------------------------------"
WScript.Echo "[!] Done searching: " & strDomain
Wscript.Quit(0)

Function PingStatus(strComputer, strDomain)
	On Error Resume Next
	strWorkstation    = strDomain '"."
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strWorkstation & "\root\cimv2")
	Set colPings      = objWMIService.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" & strComputer & "'")
	For Each objPing in colPings
		Select Case objPing.StatusCode
			Case 0 PingStatus = "Success"
			Case 11001 PingStatus = "Status code 11001 - Buffer Too Small"
			Case 11002 PingStatus = "Status code 11002 - Destination Net Unreachable"
			Case 11003 PingStatus = "Status code 11003 - Destination Host Unreachable"
			Case 11004 PingStatus = "Status code 11004 - Destination Protocol Unreachable"
			Case 11005 PingStatus = "Status code 11005 - Destination Port Unreachable"
			Case 11006 PingStatus = "Status code 11006 - No Resources"
			Case 11007 PingStatus = "Status code 11007 - Bad Option"
			Case 11008 PingStatus = "Status code 11008 - Hardware Error"
			Case 11009 PingStatus = "Status code 11009 - Packet Too Big"
			Case 11010 PingStatus = "Status code 11010 - Request Timed Out"
			Case 11011 PingStatus = "Status code 11011 - Bad Request"
			Case 11012 PingStatus = "Status code 11012 - Bad Route"
			Case 11013 PingStatus = "Status code 11013 - TimeToLive Expired Transit"
			Case 11014 PingStatus = "Status code 11014 - TimeToLive Expired Reassembly"
			Case 11015 PingStatus = "Status code 11015 - Parameter Problem"
			Case 11016 PingStatus = "Status code 11016 - Source Quench"
			Case 11017 PingStatus = "Status code 11017 - Option Too Big"
			Case 11018 PingStatus = "Status code 11018 - Bad Destination"
			Case 11032 PingStatus = "Status code 11032 - Negotiating IPSEC"
			Case 11050 PingStatus = "Status code 11050 - General Failure"
			Case Else PingStatus = "Status code " & objPing.StatusCode & " - Unable to determine cause of failure."
		End Select
	Next
End Function

wscript.quit(0)
'Pause for visual inspection
'wscript.stdin.ReadLine
'objNetwork = Nothing
'objFile = Nothing
'objErrorFSO = Nothing
'objErrorFile = Nothing
'objFSO = Nothing
'Fso = Nothing
'InputFile = Nothing
