'Script to list all servers in this Domain, make sure they respond to a ping and then write to a server.txt file

Dim objFSO, objFile

Set objFso = CreateObject("Scripting.FileSystemObject")

outFile="servers.txt"
Set objFile = objfso.CreateTextFile(outFile,True)


getServers


Function getServers() 
'Get all activa servers in the domain and write to a file 

Dim objRootDSE, strDNSDomain, adoConnection, adoCommand, strQuery
Dim adoRecordset, strComputerDN, strBase, strFilter, strAttributes

' Determine DNS domain name from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory for all computers.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Filter on computer objects with server operating system.
strFilter = "(&(objectCategory=computer)(operatingSystem=*server*))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "distinguishedName"

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 500
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute

' Enumerate computer objects with server operating systems.
Do Until adoRecordset.EOF
    strComputerDN = adoRecordset.Fields("distinguishedName").Value
    	strComputer = split(strComputerDN,"=")
	strComputerName = split(strComputer(1),",")
	Wscript.Echo strComputerName(0)
	
    strPingStatus = PingStatus(strComputerName(0))

	if strPingStatus = "Success" Then
			Wscript.Echo "[i] Successful pinging : " & strComputerName(0)
			objFile.writeline strComputerName(0)
	Else
			Wscript.Echo "[-] Failure pinging " & strComputerName(0) & ": " & strPingStatus			
	End If	

	
    adoRecordset.MoveNext
Loop

' Clean up.
objFile.Close
adoRecordset.Close
adoConnection.Close
End function

Function PingStatus(strComputer)
	On Error Resume Next
	'strWorkstation    = strDomain '"."
	strWorkstation    = "."
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
