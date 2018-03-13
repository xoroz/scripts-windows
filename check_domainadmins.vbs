' Check list of all Domain Admin users active 
' Compare with last check, if NEW USERS are added output CRITICAL (used with nagios/centreon) 
' note: removing users will not ALERT
' By Felipe Ferreira 03/2018

Option Explicit 
Const ForAppending = 8

Dim intcount : intcount = 0 
Dim objFSO , objFile , objTextFile 
Dim strShare, strFile, strFileLast, blnVerbose, blnLogging, blnFirstRun, strGroup, intExitCode, strOut, intCountMiss
Dim aUsers() ' array to load users of last run
blnFirstRun = 0
intExitCode = 0 

' EDIT HERE 
strGroup = "Domain Admins"
strFile = "c:\domainadmins.txt"


blnVerbose=0 
blnLogging=1 
strFileLast = strFile & ".last"


sub checkgroup(StrGroupName) 
	Dim ObjRootDSE, ObjConn, ObjRS, ObjCustom 
	Dim StrDomainName, StrSQL 
	Dim StrGroupDN, StrEmptySpace 
	
	Set ObjRootDSE = GetObject("LDAP://RootDSE") 
	StrDomainName = Trim(ObjRootDSE.Get("DefaultNamingContext")) 
	Set ObjRootDSE = Nothing 
	 
	' -- Mention any AD Group Name Here. Also works for Domain Admins, Enterprise Admins etc. 
	StrSQL = "Select ADsPath From 'LDAP://" & StrDomainName & "' Where ObjectCategory = 'Group' AND Name = '" & StrGroupName & "'" 
	 
	Set ObjConn = CreateObject("ADODB.Connection") 
	ObjConn.Provider = "ADsDSOObject":    ObjConn.Open "Active Directory Provider" 
	Set ObjRS = CreateObject("ADODB.Recordset") 
	ObjRS.Open StrSQL, ObjConn 
	If ObjRS.EOF Then 
		'pt VbCrLf & "This Group: " & StrGroupName & " does not exist in Active Directory" 
		Exit Sub 
	End If 
	If Not ObjRS.EOF Then     
		pt vbNullString 
		ObjRS.MoveLast:    ObjRS.MoveFirst 
		'pt "Total No of Groups Found: " & ObjRS.RecordCount 
		'pt "List of Members In " & StrGroupName & " are: " & VbCrLf 
		While Not ObjRS.EOF         
			StrGroupDN = Trim(ObjRS.Fields("ADsPath").Value) 
			Set ObjCustom = CreateObject("Scripting.Dictionary") 
			StrEmptySpace = " " 
			GetAllNestedMembers StrGroupDN, StrEmptySpace, ObjCustom 
			Set ObjCustom = Nothing 
			ObjRS.MoveNext 
		Wend 
	End If 
	ObjRS.Close:    Set ObjRS = Nothing 
	ObjConn.Close:    Set ObjConn = Nothing 
	'if (intcount > 0 ) then 
    '		pt "Total of " & intcount & " active users in group " & StrGroupName
    'end if 
end sub 


Function GetAllNestedMembers (StrGroupADsPath, StrEmptySpace, ObjCustom) 
    Dim ObjGroup, ObjMember 
    Set ObjGroup = GetObject(StrGroupADsPath) 
    For Each ObjMember In ObjGroup.Members         
        If Strcomp(Trim(ObjMember.Class), "Group", vbTextCompare) = 0 Then 		
            If ObjCustom.Exists(ObjMember.ADsPath) Then    
                pt StrEmptySpace & " -- Already Checked Group-Member " & "(Stopping Here To Escape Loop)" 
            Else 		
                ObjCustom.Add ObjMember.ADsPath, 1     
                GetFromHere ObjMember.ADsPath, StrEmptySpace & " ", ObjCustom 
            End If 
		else 				
				If Strcomp(Trim(ObjMember.AccountDisabled), "False", vbTextCompare) = 0 Then 
					pt ObjMember.sAMAccountName
					checkifuser(ObjMember.sAMAccountName)
					'pt Trim(ObjMember.CN) & " --- " & Trim(ObjMember.DisplayName) & " (" & Trim(ObjMember.Class) & ")" 
					intcount = intcount + 1 
				end if 
        End If 
    Next 
End Function 
 
Sub GetFromHere(StrGroupADsPath, StrEmptySpace, ObjCustom) 
    Dim ObjThisGroup, ObjThisMember 
    Set ObjThisGroup = GetObject(StrGroupADsPath)    
    For Each ObjThisMember In ObjThisGroup.Members
		If Strcomp(Trim(ObjThisMember.AccountDisabled), "False", vbTextCompare) = 0 Then 
			'pt replace(cstr(ObjThisGroup.Name),"CN=","") & "_" & Trim(ObjThisMember.sAMAccountName)
			pt(ObjThisMember.sAMAccountName)			
			checkifuser(ObjThisMember.sAMAccountName)
			intcount = intcount + 1 
		end if 
    Next 
End Sub

sub loadlastfile()
'check if exists a version .last of the this file, open and add to an array
	Const ForReading = 1
	Dim i, objFSO, objFileLastLog
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If not (objFSO.FileExists(strFileLast)) Then
		strOut="UNKOWN - File " & strFile & ".last not found, maybe last run" 		
		blnFirstRun =  1
		exit sub 
	End If
	
	Set objFileLastLog = objFSO.OpenTextFile(strFileLast, ForReading )
	i = 0	
	
	Do Until objFileLastLog.AtEndOfStream
	  Redim Preserve aUsers(i)
	  aUsers(i) = objFileLastLog.ReadLine
	  i = i + 1
	Loop
	
	objFileLastLog.Close
end sub 

sub checkifuser(strUser)
'SEARCH ARRAY 
	Dim j, objFile, strWriteTxt, blnFound
	blnFound = 0 		
	if (blnFirstRun = 1) then
	'nothing to compare too 
		exit sub
	end if 	
	For j=Lbound(aUsers) to Ubound(aUsers)
		'If InStr(aUsers(j), strUser) <> 0 Then
		if (aUsers(j) = strUser) then 
			blnFound = 1 
			'exit for 
		End If
	Next
	if (blnFound = 0) and (blnFirstRun = 0 ) then 
		strOut="CRITICAL - Users " & strUser & " was not in the Group before!"
		intCountMiss = intCountMiss + 1 
		intExitCode=2 
	end if 
end sub 

sub pt(txtmsg)
	'on error resume next 
	Dim objFSOL,objFileLog
	if blnVerbose = 1 then
		wscript.echo txtmsg
	end if
	if blnLogging = 1 then
		Set objFSOL = CreateObject("Scripting.FileSystemObject")
		Set objFileLog = objFSOL.OpenTextFile(strFile, ForAppending, True )
			objFileLog.Writeline txtmsg 
			objFileLog.close 
	end if 
end sub 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'1st thing move existing file to .last so we can compare with it.
dim filesys : set filesys=CreateObject("Scripting.FileSystemObject") 
If (filesys.FileExists(strFile)) Then 
	filesys.CopyFile strFile, strFileLast
	filesys.DeleteFile strFile	
	loadlastfile()
else
	wscript.echo "UNKOWN - Could not find last run file :" & strFile
    blnFirstRun = 1 
	intExitCode = 3 
End If 


checkgroup(StrGroup)
    if ( blnFirstRun = 0 ) and ( intExitCode = 0 ) then 
		wscript.echo "OK - Found " & intcount & " users on the group " & strGroup & "|users=" & intcount 
		wscript.quit(intExitCode)
	end if 
	
intCount = intCount - intCountMiss	
wscript.echo strOut & "|users=" & intcount	
wscript.quit(intExitCode)	
	
