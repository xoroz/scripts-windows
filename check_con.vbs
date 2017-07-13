'VBscript to check number of ESTABLISHED CONNECTIONS GROUP BY OFFICE IPS
'run netstat command and parse the results 
'Felipe Ferreira 13/07/2017 - Monitor Number of ESTABLISHED connections

'Declare constant Variables 
Const ForReading = 1 : Const ForWriting = 2 : Const ForAppending = 8

Dim intShowCmd,oShell,strLogFile,verbose,strMSG,intWARN,intCRIT
Dim intCountEsta

Set objShell = CreateObject("WScript.Shell")
Set colEnvironment = objShell.Environment("PROCESS")
objPath = colEnvironment("temp")
Set objShel = nothing
Set colEnvironment = nothing 

verbose=1
strLogFile=objPath & "\netstat.txt"
intShowCmd=0


intPort=":" & WScript.Arguments.Item(0)
intWARN=WScript.Arguments.Item(1)
intCRIT=WScript.Arguments.Item(2)


'--------- MAIN ----------
check_con
parsefile

strMSG="There are " & intCountEsta & " connections Established on port " & WScript.Arguments.Item(0) & " |connections=" & intCountEsta

if (intCountEsta > intCRIT) then
 pt "CRITICAL - " & strMSG
 wscript.quit 2
elseif (intCountEsta > intWARN) then
 pt "WARNING - " & strMSG
 wscript.quit 1
else
 pt "OK - " & strMSG
 wscript.quit 0
end if  

'--------- FUNCTIONS ----------
 
function check_con()
	Dim strCmd,R
	Set oShell = WScript.CreateObject ("WScript.Shell")	
	strCmd="cmd /c netstat -an | findstr /n ESTABLISHED |find /c " & chr(34) & intPort & chr(34) & " > " & strLogFile
	'strCmdDetail="cmd /c netstat -an | findstr ESTABLISHED |findstr  " & chr(34) & intPort & chr(34) & " > " & strLogFileDetail
	'pt strCmd
	R = oshell.Run(strCmd,intShowCmd, true)		
	'pt R 
end function 

function parsefile()
on error resume next 
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")		
	if objfso.FileExists(strLogFile)   Then
		Set objFile = objFSO.OpenTextFile(strLogFile, ForReading)
		intCountEsta = trim(objFile.ReadLine)
	end if 	
end function

sub pt(txtmsg)
	on error resume next 
	if verbose = 1 then
		'script.echo txtmsg
		WScript.stdout.WriteLine(txtmsg)
	end if	
end sub 


