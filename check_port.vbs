'VBScript to check a remote port 
'By Felipe Ferreira Fev 2017
'
'REQUIRES: nc.exe (netcat)

Dim strcmd,intIP,intPort,args,intTmeout 
args = WScript.Arguments.Count
If args = 2 then
  intIP=WScript.Arguments.Item(0)
  intPort=WScript.Arguments.Item(1)  
else
  wscript.echo "UNKOWN - Missing arguments: IP PORT TIMEOUT(sec)"  
  wscript.quit 3
end If

intTimeout=5 'kind of useless

strcmd = "cmd /C ""nc.exe -z -v -w " & intTimeout & " " & intIP & " " & intPort & """"

'---------------------------- FUNTCION/SUB ------------------
Sub PortMonitor(strCommand)
  dtmStart = Timer()
  'wscript.echo "Running " & strCommand
  Set StdOut = WScript.StdOut
  Set objShell = CreateObject("WScript.Shell")  
  strResult = objShell.run(strCommand , 0, true)
  'wscript.echo "RESULT= " & strResult
  dtmEnd = Timer()  
  dtmDiff = FormatNumber(dtmEnd - dtmStart, 2)

  if strResult = 0 then
     Wscript.echo "OK - Connected to " & intIP & " " & intPort & " in " & dtmDiff & " ms |connect_time=" & dtmDiff
	   wscript.quit 0
  else 
     dtmDiff="0,00"
	   Wscript.echo "CRITICAL - Could not connect to " & intIP & " " & intPort & "|connect_time=" & dtmDiff
	   wscript.quit 2
  End if
end Sub
'----------------------------- MAIN ------------------------------------

Call PortMonitor(strcmd)
