'list all local users of Adminisitrator and Remote Desktop,id,desctiption,enable/disable, last_login(if local)
'by Felipe Ferreira Jan 2017

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
Set InputFile    = fso.OpenTextFile("servers_admin.txt")

strErrorLogFile  = "Errors-ListAdmins-" & Year(Date) & Month(Date) & Day(Date) & ".log"
Set objErrorFSO  = CreateObject("Scripting.FileSystemObject")
Set objErrorFile = objErrorFSO.OpenTextFile(strErrorLogFile, Append, True)

strLogFile       = "LocalAdminRDPUsers-" & Year(Date) & Month(Date) & Day(Date) & ".log"
WScript.Echo "[i] Beginning Search"

Do While Not (InputFile.AtEndOfStream)
	On Error Resume Next
	Set objNetwork = Wscript.CreateObject("Wscript.Network")
	strDomain      = objNetwork.ComputerName	
	
	Set objFSO  = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(strLogFile, Append, True)

	strComputer   = InputFile.ReadLine
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
		
    ' SID=S-1-5-32-544, local admins
	' SID=S-1-5-32-555, Remote Desktop 	
	Set colItems = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\" & strDomain & "\root\cimv2"). _
	ExecQuery ("Select * From Win32_Group Where Domain = '" & strDomain & "' AND LocalAccount = TRUE AND SID = 'S-1-5-32-544' OR SID = 'S-1-5-32-555' ")
	
'LOOP THRU GROUPS	
	For Each objItem in colItems		
		strGroup = objItem.Name				
		WScript.Echo "-----------------------------------------------------"
		WScript.Echo "[i] Enumerating Group: " & strGroup
		Set objGroup  = GetObject("WinNT://" & strComputer & "/" & strGroup & ",group")		
		
		strResult     = strComputer & ";" & strGroup & ";"	

		If IsObject(objGroup) Then
			WScript.Echo "[*] Binded to " & strComputer & " Successfully."
		Else
			WScript.Echo "[-] Failed to bind to " & strComputer & "."
			objErrorFile.Write strComputer & "; 'Failed bind' " & "Error " & "(" & Err.Number & "):" & Err.Description & _
			vbTab & Err.Source & vbCRLF
			Exit For
		End If
       
		For Each objUser In objGroup.Members

			If Err.Number <> 0 Then
				WScript.Echo "Error " & "(" & Err.Number & "):" & Err.Description & vbTab & Err.Source
				Exit For
			End If

			

			strUserPath = objUser.aDSPath
			strUserPath = Replace(strUserPath, "WinNT://", "")
			Pos         = InStr (strUserPath, strComputer)

			If Pos > 0 Then
				strUserPath = Mid(strUserPath, Pos)
			End If
			WScript.Echo "[i] Users found: " & strUserPath
			'wscript.echo "UserPath: " & strUserPath
' should only get locallastlogin for Local users!			
			if instr(objUser.aDSPath,strComputer) then			
			 On Error Resume Next
'get last login 
			 Set objUserLocal = GetObject("WinNT://" & strComputer & "/" & objUser.Name)			 
			 intLastLogin = objUserLocal.LastLogin			 
			 If Err = 0 Then
		 		WScript.Echo "[i] Local Last Login: " & intLastLogin				
			 Else			   
			    intLastLogin = ""
				WScript.Echo "[i] X Local Last Login: " & intLastLogin
			 End If
			 On Error GoTo 0
'now should get Local User status and Description using WMI 
			 WQL="SELECT * FROM Win32_UserAccount Where LocalAccount = ""True"" AND Name = " & Quote & ObjUser.Name & Quote & ""
			 'wscript.echo WQL
'!!Win2000 and Win2012 breaks here			 
			 'Set objLWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
			 Set objLWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\" & strComputer & "\root\cimv2")
			 Set colLocalItems = objLWMIService.ExecQuery(WQL)
			 For Each objLocalItem in colLocalItems 				
				strDescription = objLocalItem.Description
				strStatus = objLocalItem.Status
			 Next
 			 WScript.Echo "[i] Local User Status: " & strStatus
			 WScript.Echo "[i] Local User Description: " & strDescription
			else			 
			 intLastLogin = ""			 
			 strDescription = ""
			 strStatus = ""
			 
			end if		
			objFile.Write strResult & ";" & strUserPath & ";" & strStatus & ";" & intLastLogin & ";" & strDescription & vbCrLf			
		Next

		WScript.Echo "[i] Done checking Group: " & strGroup
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
