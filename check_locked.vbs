' FindLockedOutUsers.vbs
' VBScript program to find locked out users.
'
' ----------------------------------------------------------------------
' Copyright (c) 2008-2010 Richard L. Mueller
' Hilltop Lab web site - http://www.rlmueller.net
' Version 1.0 - January 30, 2008
' Version 1.1 - April 8, 2008 - Bug fix, misspelled objectClass.
' Version 1.2 - November 6, 2010 - No need to set objects to Nothing.
'
' You have a royalty-free right to use, modify, reproduce, and
' distribute this script file in any way you find useful, provided that
' you agree that the copyright owner above has no warranty, obligations,
' or liability for such use.

'edited by: Felipe Ferreira 2018 (adapted for nagios/centreon output )

Option Explicit

Dim objRootDSE, strDNSDomain, objShell, lngBiasKey, lngBias, k
Dim objDomain, objDuration, lngHigh, lngLow, lngDuration
Dim adoCommand, adoConnection, adoRecordset
Dim strBase, strFilter, strAttributes, strQuery, strUser
Dim strUserDN, dtmLockOut, lngSeconds, str64Bit, intCount , intMaxCount

strUser = "" 
intMaxCount = 3 

' Retrieve DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Obtain local Time Zone bias from local machine registry.
' This bias changes with Daylight Savings Time.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngBias = lngBias + (lngBiasKey(k) * 256^k)
    Next
End If
Set objShell = Nothing

' Retrieve domain lockout duration policy.
Set objDomain = GetObject("LDAP://" & strDNSDomain)
Set objDuration = objDomain.lockoutDuration
lngHigh = objDuration.HighPart
lngLow = objDuration.LowPart
If (lngHigh = 0 And lngLow = 0) Then
    ' There is no domain lockout duration policy.
    ' Locked out accounts remain locked out until reset.
    ' Any user with value of lockoutTime greater than 0
    ' is locked out.
    str64Bit = "1"
Else
    ' Account for error in IADsLargeInteger property methods.
    If (lngLow < 0) Then
        lngHigh = lngHigh + 1
    End If
    ' Convert to minutes.
    lngDuration = lngHigh * (2^32) + lngLow
    lngDuration = -lngDuration/(60 * 10000000)

    ' Determine critical time in the past. Any accounts
    ' locked out after this time will still be locked out,
    ' unless the account has been reset (in which case the
    ' value of the lockoutTime attribute will be 0).
    ' Any accounts locked out before this time will no
    ' longer be locked out.
    ' Trap error if lockoutDuration -1 (2^63 - 1).
    On Error Resume Next
    dtmLockout = DateAdd("n", -lngDuration, Now())
    If (Err.Number <> 0) Then
        On Error GoTo 0
        ' There is no domain lockout duration policy.
        ' Locked out accounts remain locked out until reset.
        ' Any user with value of lockoutTime greater than 0
        ' is locked out.
        str64Bit = "1"
    Else
        On Error GoTo 0
        ' Convert to UTC.
        dtmLockout = DateAdd("n", lngBias, dtmLockout)

        ' Find number of seconds since 1/1/1601.
        lngSeconds = DateDiff("s", #1/1/1601#, dtmLockout)

        ' Convert to 100-nanosecond intervals. This is the
        ' equivalent Integer8 value (for this time zone).
        str64Bit = CStr(lngSeconds) & "0000000"
    End If
End If

' Use ADO to search Active Directory.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open = "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Search entire domain.
strBase = "<LDAP://" & strDNSDomain & ">"
' Filter on all user objects that are still locked out.
strFilter = "(&(objectCategory=person)(objectClass=user)(lockoutTime>=" & str64Bit & "))"
' Comma delimited list of attribute values to retrieve.
strAttributes = "distinguishedName,sAMAccountName"
' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Run the query.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 60
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute

' Enumerate the resulting recordset and display
' names of all locked out users.
intCount=0 
Do Until adoRecordset.EOF
	intCount = intCount + 1 
    strUserDN = adoRecordset.Fields("distinguishedName").Value
	strUser = strUser & ", " & adoRecordset.Fields("sAMAccountName").Value
    'Wscript.Echo strUserDN
	
    adoRecordset.MoveNext
Loop

' Clean up.
adoRecordset.Close
adoConnection.Close

if ( intCount = 0 ) then 
 wscript.echo "OK - Found " & intCount & " locked accounts|accountslocked="& intCount
 wscript.quit(0)
elseif ( intCount > intMaxCount ) then 
 wscript.echo "CRITICAL - Found " & intCount & " locked accounts: " & strUser & "|accountslocked="& intCount
 wscript.quit(2)
elseif ( intCount > 0 ) then 
 wscript.echo "WARNING - Found " & intCount & " locked accounts: " & strUser & "|accountslocked="& intCount
 wscript.quit(3)
else
 wscript.echo "OK - Found " & intCount & " locked accounts: " & strUser & "|accountslocked="& intCount
 wscript.quit(0) 
end if 
