' Count files recursive inside a folder
' Felipe Ferreira 10/2017 www.felipeferreira.net 

'Option Explicit
Dim argcountcommand
Dim arg(2)
Dim warn     'get from arg -w  
Dim crit    'get from arg -c 
Dim path        'get from arg -p
Dim debug : debug = 0 
Dim oFile
Dim intError
dim strScriptFile : strScriptFile = WScript.ScriptFullname
Dim intCountT : intCountT=0

debug =  0    ' SET 0 FOR SILENT MODE!
dim fso: set fso = CreateObject("Scripting.FileSystemObject")

GetArgs()
if argcountcommand = 0 then
    help()
elseif ((UCase(wscript.arguments(0))="-H") Or (UCase(wscript.arguments(0))="--HELP")) then
    help()
elseif(1 < argcountcommand < 5) then
    path =  GetOneArg("-p")  
	warn = GetOneArg("-w")    
	crit = GetOneArg("-c")  
end if


Call GetFilesCount(path)
strmsg="Found Total of " & intCountT & " files in " & path & "|files=" & intCountT

if ( intCountT > crit )then
 wscript.echo "CRITICAL - " & strmsg 
 wscript.quit 2
elseif ( intCountT > warn )then
 wscript.echo "WARNING - " & strmsg 
 wscript.quit 1 
else 
 wscript.echo "OK - " & strmsg 
 wscript.quit 0
end if 

 

Sub GetFilesCount(strpath)
	'path = chr(34) & path & chr(34)
    'wscript.echo strpath 
	dim folder: set folder = fso.getFolder(strpath)
	Set colFiles = folder.Files	
		 intCount =+ folder.files.Count
		 For Each objFile in colFiles 
			'Wscript.Echo objFile.Name   
			intCountT=+ 1 + intCountT
		Next
'Go recurstive in each folder 		
		For Each Subfolder in Folder.SubFolders
			GetFilesCount(Subfolder)
		Next		
	'wscript.echo "Found " & intCount & " files in " & strpath 
end sub  

Function Help()
'Prints out help    
        Dim str
        str="Check if a file exists inside a folder."&vbCrlF
		str="Also part of the filename and only for files modified today."&vbCrlF
        str=str&"cscript "& strScriptFile &" -p Path -f filename "&vbCrlF
        str=str&"cscript "& "cscript check_filesize.vbs -p c:\ -f vtapi.dll -w20 -c 30"&vbCrlF
        str=str&vbCrlF
        str=str&"-h [--help]        "&vbCrlF
        str=str&"-p path            "&vbCrlF  
        str=str&"-w <warning>		"&vbCrlF          
		str=str&"-c <critical>      "&vbCrlF          
        str=str&vbCrlF
        str=str&"By Felipe Ferreira October 2017, version 2.0." & vbCrlF
        wscript.echo str
        wscript.quit        
End Function


'@@@@@@@@@@@HANDLES THE ARGUMENTS@@@@@@@@@@@@@@@

Function GetArgs()
'Get ALL arguments passed to the script
    On Error Resume Next        
    Dim i       
    argcountcommand=WScript.Arguments.Count     
    for i=0 to argcountcommand-1
        arg(i)=WScript.Arguments(i)
        'pt i & " - " & arg(i)
    next        
End Function
Function GetOneArg(strName)
    On Error Resume Next
    Dim i
    for i=0 to argcountcommand-1
        if (Ucase(arg(i))=Ucase(strName)) then
            GetOneArg=arg(i+1)
            Exit Function
        end if
    next        
End Function

 
 
function pt(txt)
if debug = 1 then
    wscript.echo txt
end if
end function
