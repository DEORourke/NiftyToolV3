'	must be called with cscript instead of wscript in order for stdout to function correctly.
'	Device info will be passed in individual arguements to the script. An array is then made from the arguements.
'	For Each statement passes each arguement to a subroutine to device and do the ping test.
Dim arrInput()
Set objShell = CreateObject("WScript.Shell")
Set objStdOut = WScript.StdOut


redim arrInput (Wscript.Arguments.Count -1)
for I = 0 to Wscript.Arguments.Count -1
	arrInput(I) = Wscript.Arguments(I)
Next

objStdOut.Write("Performing ping test for " & UBound(arrInput) + 1 & " devices.")
objStdOut.WriteBlankLines(2)

for each device in arrInput
	pingDevice(device)
Next

Private Function pingDevice(device)
	dim avgPing
	dim deviceAddress
	dim deviceName
	dim failure
	
'		Parse device name from device IP and do the ping
	deviceName = Right(device,Len(device) - InStr(device, "-"))
	deviceAddress = Left(device, InStr(device, "-") -1 )

	objStdOut.Write("Pinging " & deviceName & "... ")
	Set objPing = objShell.Exec("ping -n 3 " & deviceAddress )
	strPingResults = objPing.StdOut.Readall
	avgPing = Right(strPingResults, Len(strPingResults) - InstrRev(strPingResults,"Average") - 9 )
	failure = InStr(strPingResults,"100% loss")	' searching for this string to determine if the ping attempt completly failed.

	If failure = 0 Then
		objStdOut.Write(deviceName & " appears online. Avg ping time: " & avgPing)
	ElseIf failure > 0 Then
		objStdOut.Write(deviceName & " is not responding to ping request." & vbNewLine)
	Else
		objStdOut.Write("Something went wrong when trying to ping " & deviceName & ".")
	End IF

End Function

objStdOut.WriteBlankLines(1)